import * as pdfjs from "./vendor/pdf.min.mjs";
pdfjs.GlobalWorkerOptions.workerSrc = new URL(
  "./vendor/pdf.worker.min.mjs",
  import.meta.url,
).href;
const pdfOptions = {
  isEvalSupported: false,
  cMapUrl: new URL("./vendor/cmaps/", import.meta.url).href,
  cMapPacked: true,
  standardFontDataUrl: new URL("./vendor/standard_fonts/", import.meta.url)
    .href,
};
const { PDFDocument, degrees, PDFName, PDFRawStream, PDFNumber } = PDFLib;
const $ = (id) => document.getElementById(id),
  MB = 1000000;
const LIMIT = { files: 30, pages: 100, bytes: 200 * MB, pixels: 40_000_000 };
const colors = ["#4d78ac", "#d88741", "#578d7c", "#8c76ac", "#aa6880"];
let sources = new Map(),
  pages = [],
  undo = [],
  redo = [],
  outputs = [],
  zipBlob = null;
let busy = false,
  cancelled = false,
  revision = 0,
  renderVersion = 0,
  selectedOutput = 0,
  comparison = null,
  viewVersion = 0;
const thumbs = new Map(),
  urls = new Set();
const message = (text = "") => ($("notice").textContent = text);
const size = (n) =>
  n < 1000000
    ? (n / 1000).toFixed(1) + " KB"
    : (n / 1000000).toFixed(2) + " MB";
const tick = () => new Promise((r) => setTimeout(r, 0));
const assertActive = () => {
  if (cancelled) throw new Error("작업을 취소했습니다.");
};
const safeName = (n) =>
  (n || "문서")
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, "_")
    .replace(/\.pdf$/i, "")
    .replace(/[. ]+$/, "")
    .slice(0, 80) || "문서";
function el(tag, text, cls) {
  const e = document.createElement(tag);
  if (text !== undefined) e.textContent = text;
  if (cls) e.className = cls;
  return e;
}
function button(text, fn) {
  const b = el("button", text);
  b.onclick = fn;
  return b;
}
function post(type, extra = {}) {
  if (window.parent !== window)
    window.parent.postMessage(
      { isStreamlitMessage: true, type, ...extra },
      "*",
    );
}
post("streamlit:componentReady", { apiVersion: 1 });
let frameHeight = 0;
new ResizeObserver(() => {
  const h = Math.ceil(document.body.getBoundingClientRect().height) + 4;
  if (h !== frameHeight) {
    frameHeight = h;
    post("streamlit:setFrameHeight", { height: h });
  }
}).observe(document.body);
function invalidate() {
  revision++;
  clearResults();
  $("dirty").textContent =
    "다시 생성 필요 · 현재 편집 내용으로 결과를 만들어 주세요.";
}
function clearResults() {
  for (const x of outputs) x.viewer?.destroy();
  outputs = [];
  zipBlob = null;
  for (const u of urls) URL.revokeObjectURL(u);
  urls.clear();
  $("results").replaceChildren();
  $("reviewPages").replaceChildren();
}
function commit(fn) {
  if (busy) return;
  undo.push(structuredClone(pages));
  if (undo.length > 40) undo.shift();
  redo = [];
  fn();
  invalidate();
  render();
}
function history(from, to) {
  if (busy || !from.length) return;
  to.push(structuredClone(pages));
  pages = from.pop();
  invalidate();
  render();
}
function settings() {
  const v = $("target").value;
  const target = v === "custom" ? Number($("custom").value) : Number(v);
  if (
    !Number.isFinite(target) ||
    target < 0 ||
    (v === "custom" && (target < 0.1 || target > 200))
  )
    throw new Error("목표 용량을 0.1~200MB 사이로 입력해 주세요.");
  return {
    mode: $("mode").value,
    name: safeName($("filename").value),
    target: Math.floor(target * MB),
    quality: $("quality").value,
    a4: $("size").value === "a4",
    raster: $("raster").checked,
  };
}
function setBusy(value, label = "") {
  busy = value;
  $("editor").inert = value;
  $("review").inert = value;
  $("reset").disabled = value;
  $("busy").hidden = !value;
  $("progress").textContent = label;
  cancelled = false;
}
function progress(t) {
  $("progress").textContent = t;
  assertActive();
}
async function jpeg(canvas, quality) {
  const blob = await new Promise((resolve, reject) =>
    canvas.toBlob(
      (b) =>
        b ? resolve(b) : reject(new Error("이미지 압축에 실패했습니다.")),
      "image/jpeg",
      quality,
    ),
  );
  return new Uint8Array(await blob.arrayBuffer());
}
async function bitmap(blob) {
  const b = await createImageBitmap(blob, { imageOrientation: "from-image" });
  if (!b.width || b.width * b.height > LIMIT.pixels) {
    b.close();
    throw new Error("이미지는 한 장당 4,000만 화소 이하로 줄여 주세요.");
  }
  return b;
}
function canvasFor(w, h) {
  const c = document.createElement("canvas");
  c.width = Math.max(1, Math.round(w));
  c.height = Math.max(1, Math.round(h));
  return c;
}
async function imageCanvas(source, edge = 2600, rotation = 0) {
  const b = await bitmap(source.file);
  const turn = rotation % 180 !== 0;
  const w = turn ? b.height : b.width,
    h = turn ? b.width : b.height;
  const scale = Math.min(1, edge / Math.max(w, h));
  const c = canvasFor(w * scale, h * scale);
  const ctx = c.getContext("2d");
  ctx.fillStyle = "white";
  ctx.fillRect(0, 0, c.width, c.height);
  ctx.translate(c.width / 2, c.height / 2);
  ctx.rotate((rotation * Math.PI) / 180);
  ctx.drawImage(
    b,
    (-b.width * scale) / 2,
    (-b.height * scale) / 2,
    b.width * scale,
    b.height * scale,
  );
  b.close();
  return c;
}
async function getViewer(source) {
  if (!source.viewer)
    source.viewer = await pdfjs.getDocument({
      data: source.bytes.slice(),
      ...pdfOptions,
    }).promise;
  return source.viewer;
}
async function renderPdf(viewer, index, edge, rotation = 0) {
  const p = await viewer.getPage(index + 1);
  const v = p.getViewport({ scale: 1, rotation: (p.rotate + rotation) % 360 });
  const scale = Math.min(edge / Math.max(v.width, v.height), 4);
  const vp = p.getViewport({ scale, rotation: (p.rotate + rotation) % 360 });
  const c = canvasFor(vp.width, vp.height);
  await p.render({
    canvasContext: c.getContext("2d"),
    viewport: vp,
    background: "white",
  }).promise;
  p.cleanup();
  return c;
}
async function pageCanvas(page, edge = 450, original = true) {
  const src = sources.get(page.source);
  if (src.type === "image") return imageCanvas(src, edge, page.rotation);
  return renderPdf(await getViewer(src), page.index, edge, page.rotation);
}
async function hash(bytes) {
  return Array.from(
    new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)),
  )
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
}
async function addFiles(files) {
  if (busy) return;
  setBusy(true, "파일을 확인하고 있습니다…");
  $("errors").textContent = "";
  const failures = [];
  let changed = false;
  const before = structuredClone(pages);
  try {
    for (const file of files) {
      assertActive();
      let src = null;
      try {
        progress(file.name + " 확인 중…");
        if (sources.size >= LIMIT.files)
          throw new Error("파일은 최대 30개까지 추가할 수 있습니다.");
        if (
          [...sources.values()].reduce((n, s) => n + s.file.size, 0) +
            file.size >
          LIMIT.bytes
        )
          throw new Error("원본 전체 용량은 200MB 이하로 선택해 주세요.");
        if (!file.size) throw new Error("빈 파일입니다.");
        const bytes = new Uint8Array(await file.arrayBuffer());
        const digest = await hash(bytes);
        if (
          [...sources.values()].some((s) => s.hash === digest) &&
          !confirm(
            file.name + "은 이미 추가한 파일입니다. 한 번 더 추가할까요?",
          )
        )
          continue;
        const id = crypto.randomUUID();
        src = {
          id,
          file,
          name: file.name,
          hash: digest,
          color: colors[sources.size % colors.length],
          type: /\.pdf$/i.test(file.name) ? "pdf" : "image",
          bytes,
        };
        let count = 1;
        if (src.type === "pdf") {
          let doc;
          try {
            doc = await PDFDocument.load(bytes);
          } catch (e) {
            throw new Error(
              /encrypt|password/i.test(String(e))
                ? "암호가 있는 PDF입니다. 첫 버전에서는 암호 해제된 사본을 사용해 주세요."
                : "PDF를 읽을 수 없습니다. 손상 여부를 확인해 주세요.",
            );
          }
          const form = doc.getForm();
          if (form.getFields().some((f) => f instanceof PDFLib.PDFSignature))
            throw new Error(
              "전자서명 필드가 있는 PDF는 이 도구에서 편집할 수 없습니다.",
            );
          if (form.getFields().length) {
            form.flatten();
            src.normalized = await doc.save();
          }
          count = doc.getPageCount();
          if (!count) throw new Error("페이지가 없는 PDF입니다.");
          src.doc = doc;
          const viewer = await getViewer(src);
          if (viewer.numPages !== count)
            throw new Error("PDF 페이지 구조를 확인할 수 없습니다.");
        } else {
          if (!/\.(png|jpe?g|webp|bmp)$/i.test(file.name))
            throw new Error("PDF·JPG·PNG·WEBP·BMP 파일만 지원합니다.");
          const c = await imageCanvas(src);
          src.width = c.width;
          src.height = c.height;
          src.embedType = "jpeg";
          if (file.size < MB && Math.max(c.width, c.height) < 2600) {
            const lossless = await new Promise((r) => c.toBlob(r, "image/png"));
            src.precompressed = new Uint8Array(await lossless.arrayBuffer());
            src.embedType = "png";
          } else src.precompressed = await jpeg(c, 0.9);
          c.width = c.height = 1;
          src.bytes = null;
        }
        if (pages.length + count > LIMIT.pages)
          throw new Error("전체 페이지는 최대 100쪽까지 추가할 수 있습니다.");
        sources.set(id, src);
        for (let i = 0; i < count; i++)
          pages.push({
            id: crypto.randomUUID(),
            source: id,
            index: i,
            rotation: 0,
            selected: true,
            split: false,
          });
        changed = true;
      } catch (e) {
        if (cancelled) throw e;
        if (src?.viewer) await src.viewer.destroy();
        failures.push(file.name + ": " + e.message);
      }
      await tick();
    }
  } catch (e) {
    message(e.message);
  } finally {
    setBusy(false);
    if (changed) {
      undo.push(before);
      redo = [];
      invalidate();
    }
    render();
    $("errors").textContent = failures.join("\n");
    $("files").value = "";
  }
}
function move(id, delta) {
  const i = pages.findIndex((p) => p.id === id),
    j = i + delta;
  if (j < 0 || j >= pages.length) return;
  commit(() => {
    [pages[i], pages[j]] = [pages[j], pages[i]];
  });
}
function render() {
  const epoch = ++renderVersion;
  $("stepEdit").classList.add("active");
  $("stepReview").classList.remove("active");
  $("pages").replaceChildren();
  $("sources").replaceChildren();
  $("empty").hidden = pages.length > 0;
  $("count").textContent =
    `전체 ${pages.length}페이지 · ${pages.filter((p) => p.selected).length}페이지 선택`;
  $("generate").disabled = !pages.some((p) => p.selected);
  $("undo").disabled = !undo.length;
  $("redo").disabled = !redo.length;
  for (const src of sources.values()) {
    const own = pages.filter((p) => p.source === src.id);
    if (!own.length) continue;
    const div = el("div", undefined, "source");
    div.style.setProperty("--accent", src.color);
    const label = el("label");
    const cb = el("input");
    cb.type = "checkbox";
    cb.checked = own.every((p) => p.selected);
    cb.indeterminate = !cb.checked && own.some((p) => p.selected);
    cb.onchange = () =>
      commit(() => own.forEach((p) => (p.selected = cb.checked)));
    label.append(
      cb,
      document.createTextNode(src.name + " · " + own.length + "쪽"),
    );
    div.append(
      label,
      el(
        "small",
        src.type === "image"
          ? `원본 ${size(src.file.size)} → 준비 이미지 ${size(src.precompressed.length)}`
          : size(src.file.size),
      ),
    );
    $("sources").append(div);
  }
  let ordinal = 0;
  const pending = [];
  for (const [index, page] of pages.entries()) {
    const src = sources.get(page.source);
    const card = el(
      "article",
      undefined,
      "card" + (page.selected ? "" : " off"),
    );
    card.dataset.id = page.id;
    card.draggable = true;
    card.style.setProperty("--accent", src.color);
    card.ondragstart = (e) => {
      e.dataTransfer.setData("text/plain", page.id);
      card.classList.add("dragging");
    };
    card.ondragend = () => card.classList.remove("dragging");
    card.ondragover = (e) => e.preventDefault();
    card.ondrop = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const id = e.dataTransfer.getData("text/plain");
      if (id === page.id || !pages.some((p) => p.id === id)) return;
      commit(() => {
        const from = pages.findIndex((p) => p.id === id);
        const [x] = pages.splice(from, 1);
        pages.splice(
          pages.findIndex((p) => p.id === page.id),
          0,
          x,
        );
      });
    };
    const head = el("div", undefined, "cardHead");
    const cb = el("input");
    cb.type = "checkbox";
    cb.checked = page.selected;
    cb.setAttribute(
      "aria-label",
      src.name + " " + (page.index + 1) + "쪽 선택",
    );
    cb.onchange = () => commit(() => (page.selected = cb.checked));
    head.append(
      cb,
      el("span", page.selected ? `${++ordinal}페이지` : "제외됨"),
      el("b", "⠿", "handle"),
    );
    const thumb = el("div", undefined, "thumb");
    thumb.tabIndex = 0;
    thumb.setAttribute("role", "button");
    thumb.setAttribute("aria-label", "페이지 확대");
    thumb.onclick = () => showView(page);
    thumb.onkeydown = (e) => {
      if (e.key === "Enter") showView(page);
    };
    const img = el("img");
    img.alt = src.name + " 원본 " + (page.index + 1) + "쪽";
    thumb.append(img);
    const cacheKey = page.id + ":" + page.rotation;
    if (thumbs.has(cacheKey)) img.src = thumbs.get(cacheKey);
    else pending.push({ img, page, key: cacheKey });
    const acts = el("div", undefined, "actions");
    const prev = button("←", () => move(page.id, -1));
    prev.title = "앞으로 이동";
    prev.disabled = index === 0;
    const next = button("→", () => move(page.id, 1));
    next.title = "뒤로 이동";
    next.disabled = index === pages.length - 1;
    acts.append(
      prev,
      next,
      button("↻ 회전", () =>
        commit(() => (page.rotation = (page.rotation + 90) % 360)),
      ),
      button("확대", () => showView(page)),
    );
    card.append(
      head,
      thumb,
      el("div", src.name + " · 원본 " + (page.index + 1) + "쪽", "caption"),
      acts,
    );
    if ($("mode").value === "groups") {
      const split = button(
        page.split ? "✓ 여기서 나눔" : "＋ 여기서 나누기",
        () => commit(() => (page.split = !page.split)),
      );
      split.className = "split" + (page.split ? " on" : "");
      card.append(split);
    }
    $("pages").append(card);
  }
  (async () => {
    for (const job of pending) {
      if (epoch !== renderVersion) return;
      try {
        const c = await pageCanvas(job.page, 400);
        const data = c.toDataURL("image/jpeg", 0.8);
        c.width = c.height = 1;
        if (epoch !== renderVersion) return;
        thumbs.set(job.key, data);
        if (thumbs.size > 160) thumbs.delete(thumbs.keys().next().value);
        job.img.src = data;
      } catch {
        job.img.alt = "미리보기 실패 · 결과 생성 시 재확인";
      }
      await tick();
    }
  })();
}
function groupsFor(mode) {
  const groups = [];
  let group = [];
  for (const p of pages) {
    if (p.selected) {
      if (mode === "pages") groups.push([p]);
      else group.push(p);
    }
    if (mode === "groups" && p.split && group.length) {
      groups.push(group);
      group = [];
    }
  }
  if (group.length) groups.push(group);
  return groups;
}
async function assemble(group, opt) {
  const out = await PDFDocument.create();
  for (const row of group) {
    assertActive();
    const src = sources.get(row.source);
    if (src.type === "pdf") {
      const [copy] = await out.copyPages(src.doc, [row.index]);
      copy.setRotation(
        degrees((copy.getRotation().angle + row.rotation) % 360),
      );
      out.addPage(copy);
    } else {
      let embedded;
      if (opt.quality === "original") {
        const c = await imageCanvas(src, Infinity);
        const blob = await new Promise((r) => c.toBlob(r, "image/png"));
        embedded = await out.embedPng(await blob.arrayBuffer());
        c.width = c.height = 1;
      } else
        embedded =
          src.embedType === "png"
            ? await out.embedPng(src.precompressed)
            : await out.embedJpg(src.precompressed);
      const w = embedded.width,
        h = embedded.height;
      let pw, ph, margin;
      if (opt.a4) {
        [pw, ph] = w > h ? [842, 595] : [595, 842];
        margin = 22.7;
      } else {
        const k = 842 / Math.max(w, h);
        [pw, ph, margin] = [w * k, h * k, 0];
      }
      const p = out.addPage([pw, ph]);
      const k = Math.min((pw - 2 * margin) / w, (ph - 2 * margin) / h);
      p.drawImage(embedded, {
        x: (pw - w * k) / 2,
        y: (ph - h * k) / 2,
        width: w * k,
        height: h * k,
      });
      p.setRotation(degrees(row.rotation));
    }
    await tick();
  }
  return out.save();
}
// Recompress only uncomplicated RGB/gray JPEG XObjects. Leave masks, CMYK,
// custom Decode arrays and non-JPEG streams intact to avoid color/data corruption.
async function compressImages(bytes, edge, quality) {
  const doc = await PDFDocument.load(bytes);
  let count = 0;
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    assertActive();
    if (!(obj instanceof PDFRawStream)) continue;
    const d = obj.dict;
    const val = (n) => String(d.get(PDFName.of(n)));
    const filter = d.get(PDFName.of("Filter"));
    const ascii85 =
      filter instanceof PDFLib.PDFArray &&
      filter.size() === 2 &&
      String(filter.get(0)) === "/ASCII85Decode" &&
      String(filter.get(1)) === "/DCTDecode";
    if (
      val("Subtype") !== "/Image" ||
      (val("Filter") !== "/DCTDecode" && !ascii85) ||
      !["/DeviceRGB", "/DeviceGray"].includes(val("ColorSpace")) ||
      ["SMask", "Mask", "Decode", "DecodeParms"].some((n) =>
        d.has(PDFName.of(n)),
      )
    )
      continue;
    let b;
    try {
      let raw = obj.contents;
      if (ascii85) {
        const wrapper = d.clone(doc.context);
        wrapper.set(PDFName.of("Filter"), PDFName.of("ASCII85Decode"));
        raw = PDFLib.decodePDFRawStream(
          PDFRawStream.of(wrapper, obj.contents),
        ).decode();
      }
      b = await createImageBitmap(new Blob([raw], { type: "image/jpeg" }));
      if (b.width * b.height > LIMIT.pixels) {
        b.close();
        continue;
      }
      const k = Math.min(1, edge / Math.max(b.width, b.height));
      const c = canvasFor(b.width * k, b.height * k);
      c.getContext("2d").drawImage(b, 0, 0, c.width, c.height);
      b.close();
      b = null;
      const encoded = await jpeg(c, quality);
      if (encoded.length < obj.contents.length) {
        const dict = d.clone(doc.context);
        dict.set(PDFName.of("Width"), PDFNumber.of(c.width));
        dict.set(PDFName.of("Height"), PDFNumber.of(c.height));
        dict.set(PDFName.of("BitsPerComponent"), PDFNumber.of(8));
        dict.set(PDFName.of("ColorSpace"), PDFName.of("DeviceRGB"));
        dict.set(PDFName.of("Filter"), PDFName.of("DCTDecode"));
        dict.set(PDFName.of("Length"), PDFNumber.of(encoded.length));
        doc.context.assign(ref, PDFRawStream.of(dict, encoded));
        count++;
      }
      c.width = c.height = 1;
    } catch (e) {
      b?.close();
      if (cancelled) throw e;
    }
    await tick();
  }
  return { bytes: await doc.save(), count };
}
async function rasterize(bytes, edge, quality) {
  const viewer = await pdfjs.getDocument({ data: bytes.slice(), ...pdfOptions })
    .promise;
  try {
    const doc = await PDFDocument.create();
    for (let i = 0; i < viewer.numPages; i++) {
      assertActive();
      const pg = await viewer.getPage(i + 1);
      const v = pg.getViewport({ scale: 1 });
      const c = await renderPdf(viewer, i, edge);
      const jpg = await jpeg(c, quality);
      const embedded = await doc.embedJpg(jpg);
      const p = doc.addPage([v.width, v.height]);
      p.drawImage(embedded, { x: 0, y: 0, width: v.width, height: v.height });
      c.width = c.height = 1;
      await tick();
    }
    return doc.save();
  } finally {
    await viewer.destroy();
  }
}
async function generate() {
  if (busy) return;
  let opt;
  try {
    opt = settings();
  } catch (e) {
    message(e.message);
    return;
  }
  const groups = groupsFor(opt.mode);
  if (!groups.length) return;
  clearResults();
  setBusy(true, "PDF를 생성하고 있습니다…");
  $("errors").textContent = "";
  const prepared = [];
  let preparedZip = null;
  try {
    for (let i = 0; i < groups.length; i++) {
      progress(`${i + 1}/${groups.length} 파일 생성 중…`);
      const base = await assemble(groups[i], opt);
      let best = base,
        raster = false,
        changed = 0;
      if (
        opt.target &&
        best.length > opt.target &&
        opt.quality !== "original"
      ) {
        for (const [edge, q] of [
          [2600, 0.86],
          [2000, 0.78],
          [1600, 0.68],
          [1200, 0.58],
        ]) {
          progress(
            `${i + 1}/${groups.length} 파일의 이미지 용량을 조절하고 있습니다…`,
          );
          const next = await compressImages(base, edge, q);
          if (next.bytes.length < best.length) {
            best = next.bytes;
            changed = next.count;
          }
          if (best.length <= opt.target) break;
        }
        if (best.length > opt.target && opt.raster) {
          for (const [edge, q] of [
            [2400, 0.85],
            [1800, 0.75],
            [1400, 0.65],
          ]) {
            progress(`${i + 1}/${groups.length} 파일을 이미지화하고 있습니다…`);
            const next = await rasterize(base, edge, q);
            if (next.length < best.length) {
              best = next;
              raster = true;
            }
            if (best.length <= opt.target) break;
          }
        }
      }
      assertActive();
      prepared.push({
        name:
          opt.name +
          (groups.length > 1 ? "_" + String(i + 1).padStart(2, "0") : "") +
          ".pdf",
        bytes: best,
        baseBytes: base.length,
        pages: structuredClone(groups[i]),
        target: opt.target,
        raster,
        changed,
      });
    }
    if (prepared.length > 1) {
      progress("ZIP을 만들고 있습니다…");
      preparedZip = await makeZip(prepared);
      assertActive();
    }
    outputs = prepared;
    zipBlob = preparedZip;
    selectedOutput = 0;
    $("dirty").textContent = "";
    message("");
    $("editor").hidden = true;
    $("review").hidden = false;
    $("stepEdit").classList.remove("active");
    $("stepReview").classList.add("active");
    await renderResults();
    window.scrollTo(0, 0);
  } catch (e) {
    clearResults();
    message(
      e.message ||
        "PDF 생성에 실패했습니다. 파일을 확인하고 다시 시도해 주세요.",
    );
  } finally {
    setBusy(false);
  }
}
async function makeZip(items) {
  const zip = new JSZip();
  for (const item of items) zip.file(item.name, item.bytes);
  return zip.generateAsync({ type: "blob", compression: "STORE" });
}
function validateNames() {
  const names = outputs.map((x) => x.name.toLocaleLowerCase());
  if (new Set(names).size !== names.length)
    throw new Error("파일명이 겹칩니다. 서로 다른 이름을 입력해 주세요.");
}
async function renameResult(index, value) {
  const old = outputs[index].name;
  outputs[index].name = safeName(value) + ".pdf";
  try {
    validateNames();
    if (outputs.length > 1) {
      setBusy(true, "변경된 파일명으로 ZIP을 준비합니다…");
      zipBlob = await makeZip(outputs);
    }
    message("");
    updateDownload();
  } catch (e) {
    outputs[index].name = old;
    message(e.message);
    $("results").querySelectorAll("input")[index].value = old;
  } finally {
    setBusy(false);
  }
}
function updateDownload() {
  $("zipSize").textContent = zipBlob
    ? `${outputs.length}개 PDF · ZIP 실제 용량 ${size(zipBlob.size)}`
    : outputs.length
      ? `PDF 실제 용량 ${size(outputs[0].bytes.length)}`
      : "";
  $("download").textContent = zipBlob ? "ZIP 다운로드" : "PDF 다운로드";
}
async function resultViewer(out) {
  if (!out.viewer)
    out.viewer = await pdfjs.getDocument({
      data: out.bytes.slice(),
      ...pdfOptions,
    }).promise;
  return out.viewer;
}
async function renderResults() {
  $("results").replaceChildren();
  for (const [i, out] of outputs.entries()) {
    const row = el(
      "div",
      undefined,
      "result" + (selectedOutput === i ? " active" : ""),
    );
    const name = el("input");
    name.value = out.name;
    name.setAttribute("aria-label", `${i + 1}번 결과 파일명`);
    name.onchange = () => renameResult(i, name.value);
    const info = el("div");
    info.append(
      el("strong", `${out.pages.length}쪽 · ${size(out.bytes.length)}`),
      el("small", ` 생성 전 ${size(out.baseBytes)}`, "subtle"),
    );
    const over = out.target && out.bytes.length > out.target;
    const badge = el(
      "span",
      out.target ? (over ? "목표 초과" : "목표 충족") : "제한 없음",
      "badge" + (over ? " warn" : ""),
    );
    const show = button("순서·화질 확인", async () => {
      selectedOutput = i;
      await renderResults();
    });
    row.append(name, info, badge, show);
    $("results").append(row);
    if (out.raster)
      $("results").append(
        el(
          "p",
          "이미지화된 PDF입니다. 글자 선택·검색이 제한되므로 화질을 확인해 주세요.",
          "subtle",
        ),
      );
    if (over)
      $("results").append(
        el(
          "p",
          "목표 용량을 초과했습니다. 편집에서 추가 압축을 허용하거나, 목표를 늘리거나, 그룹을 나눌 수 있습니다. 현재 결과도 다운로드할 수 있습니다.",
          "subtle",
        ),
      );
  }
  updateDownload();
  await renderReviewPages();
}
async function renderReviewPages() {
  const out = outputs[selectedOutput];
  if (!out) return;
  const epoch = ++renderVersion;
  $("reviewPages").replaceChildren();
  const viewer = await resultViewer(out);
  for (const [i, p] of out.pages.entries()) {
    if (epoch !== renderVersion) return;
    const src = sources.get(p.source);
    const card = el("article", undefined, "card");
    card.style.setProperty("--accent", src.color);
    const thumb = el("div", undefined, "thumb");
    const img = el("img");
    img.alt = `최종 ${i + 1}페이지`;
    thumb.append(img);
    thumb.onclick = () => showView(p, out, i);
    const acts = el("div", undefined, "actions");
    acts.append(button("화질 비교", () => showView(p, out, i)));
    card.append(
      el("div", `${i + 1}페이지`, "cardHead"),
      thumb,
      el("div", `${src.name} · 원본 ${p.index + 1}쪽`, "caption"),
      acts,
    );
    $("reviewPages").append(card);
    const c = await renderPdf(viewer, i, 450);
    if (epoch !== renderVersion) return;
    img.src = c.toDataURL("image/jpeg", 0.85);
    c.width = c.height = 1;
  }
}
async function showView(page, out = null, index = 0) {
  if (busy) return;
  comparison = { page, out, index };
  $("zoom").value = "1";
  $("viewTitle").textContent =
    sources.get(page.source).name + " · 원본 " + (page.index + 1) + "쪽";
  $("resultPane").hidden = !out;
  document.querySelector(".compare").classList.toggle("solo", !out);
  document.querySelector(".mobileTabs").hidden = !out;
  $("compareHint").hidden = !out;
  $("originalPane").classList.remove("mobileHide");
  $("resultPane").classList.add("mobileHide");
  $("viewer").showModal();
  $("viewer").scrollIntoView({ block: "center" });
  await drawComparison();
}
async function drawComparison() {
  const token = ++viewVersion;
  if (!comparison) return;
  const { page, out, index } = comparison;
  const zoom = Number($("zoom").value);
  try {
    const width = Math.max(
      250,
      Math.min(
        530,
        window.innerWidth < 650
          ? window.innerWidth - 65
          : window.innerWidth / 2 - 65,
      ),
    );
    const edge = Math.min(3600, width * 1.45 * zoom * 1.6);
    const a = await pageCanvas(page, edge);
    const b = out
      ? await renderPdf(await resultViewer(out), index, edge)
      : null;
    if (token !== viewVersion) {
      a.width = a.height = 1;
      if (b) b.width = b.height = 1;
      return;
    }
    for (const [c, id] of [
      [a, "originalCanvas"],
      [b, "resultCanvas"],
    ]) {
      if (!c) continue;
      const target = $(id);
      target.width = c.width;
      target.height = c.height;
      target.style.width = width * zoom + "px";
      target.style.height = "auto";
      target.getContext("2d").drawImage(c, 0, 0);
      c.width = c.height = 1;
    }
  } catch (e) {
    message("미리보기를 만들 수 없습니다: " + e.message);
  }
}
let sync = false;
const panes = [...document.querySelectorAll(".viewport")];
for (const [i, p] of panes.entries())
  p.onscroll = () => {
    if (sync) return;
    sync = true;
    const q = panes[1 - i];
    q.scrollLeft = p.scrollLeft;
    q.scrollTop = p.scrollTop;
    requestAnimationFrame(() => (sync = false));
  };
function download() {
  if (busy || !outputs.length) return;
  const blob =
    zipBlob || new Blob([outputs[0].bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  urls.add(url);
  const a = el("a");
  a.href = url;
  a.download = zipBlob
    ? safeName($("filename").value) + ".zip"
    : outputs[0].name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => {
    URL.revokeObjectURL(url);
    urls.delete(url);
  }, 60000);
}
$("files").onchange = (e) => addFiles([...e.target.files]);
$("drop").onkeydown = (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    $("files").click();
  }
};
$("drop").ondragover = (e) => {
  e.preventDefault();
  $("drop").classList.add("drag");
};
$("drop").ondragleave = () => $("drop").classList.remove("drag");
$("drop").ondrop = (e) => {
  e.preventDefault();
  $("drop").classList.remove("drag");
  addFiles([...e.dataTransfer.files]);
};
$("selectAll").onclick = () =>
  commit(() => pages.forEach((p) => (p.selected = true)));
$("excludeAll").onclick = () =>
  commit(() => pages.forEach((p) => (p.selected = false)));
$("undo").onclick = () => history(undo, redo);
$("redo").onclick = () => history(redo, undo);
$("generate").onclick = generate;
$("cancel").onclick = () => {
  cancelled = true;
  $("progress").textContent = "현재 처리 후 취소합니다…";
};
$("back").onclick = () => {
  $("review").hidden = true;
  $("editor").hidden = false;
  render();
  window.scrollTo(0, 0);
};
$("download").onclick = download;
for (const id of [
  "mode",
  "filename",
  "target",
  "custom",
  "quality",
  "size",
  "raster",
])
  $(id).addEventListener("change", () => {
    if (busy) return;
    invalidate();
    $("customWrap").hidden = $("target").value !== "custom";
    if (id === "mode") render();
  });
$("closeView").onclick = () => {
  $("viewer").close();
  comparison = null;
  viewVersion++;
};
$("viewer").addEventListener("close", () => {
  viewVersion++;
  comparison = null;
});
$("zoom").onchange = drawComparison;
$("tabOriginal").onclick = () => {
  $("originalPane").classList.remove("mobileHide");
  $("resultPane").classList.add("mobileHide");
};
$("tabResult").onclick = () => {
  $("resultPane").classList.remove("mobileHide");
  $("originalPane").classList.add("mobileHide");
};
$("reset").onclick = async () => {
  if (busy || !confirm("현재 파일과 편집 내용을 모두 비울까요?")) return;
  clearResults();
  for (const s of sources.values()) await s.viewer?.destroy();
  sources.clear();
  pages = [];
  undo = [];
  redo = [];
  thumbs.clear();
  $("review").hidden = true;
  $("editor").hidden = false;
  $("errors").textContent = "";
  $("dirty").textContent = "";
  message("");
  render();
};
window.addEventListener("beforeunload", (e) => {
  if (pages.length) {
    e.preventDefault();
    e.returnValue = "";
  }
});
render();
$("editor").inert = false;
