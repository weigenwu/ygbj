(() => {
  "use strict";

  const $ = (selector) => document.querySelector(selector);
  const $$ = (selector) => [...document.querySelectorAll(selector)];
  const SVG_NS = "http://www.w3.org/2000/svg";
  const SUPPORTED = /\.(tif|tiff|png|jpe?g|webp)$/i;
  const STORAGE_PREFIX = "fusionmark:v1:";
  const HISTORY_LIMIT = 60;

  const el = {
    fileInput: $("#fileInput"),
    pasteButton: $("#pasteButton"),
    dropZone: $("#dropZone"),
    imageList: $("#imageList"),
    imageCount: $("#imageCount"),
    splitViewport: $("#splitViewport"),
    splitCountInput: $("#splitCountInput"),
    annotationMode: $("#annotationMode"),
    viewport: $("#viewport"),
    primaryBadge: $("#primaryBadge"),
    stage: $("#stage"),
    canvas: $("#imageCanvas"),
    overlay: $("#overlay"),
    empty: $("#emptyState"),
    loading: $("#loading"),
    hint: $("#canvasHint"),
    currentName: $("#currentName"),
    currentMeta: $("#currentMeta"),
    sizeInput: $("#sizeInput"),
    sizeOutput: $("#sizeOutput"),
    strokeInput: $("#strokeInput"),
    strokeOutput: $("#strokeOutput"),
    undo: $("#undoButton"),
    redo: $("#redoButton"),
    delete: $("#deleteButton"),
    clear: $("#clearButton"),
    copy: $("#copyButton"),
    export: $("#exportButton"),
    exportNext: $("#exportNextButton"),
    zoomIn: $("#zoomInButton"),
    zoomOut: $("#zoomOutButton"),
    fit: $("#fitButton"),
    zoomOutput: $("#zoomOutput"),
    toast: $("#toast")
  };

  const state = {
    items: [],
    current: -1,
    loadedId: null,
    loadToken: 0,
    tool: "arrow",
    pathology: false,
    color: "#ffffff",
    size: 5,
    stroke: 5,
    scale: 1,
    panX: 0,
    panY: 0,
    fitScale: 1,
    selectedId: null,
    interaction: null,
    splitCount: 1,
    compareIndices: [-1, -1, -1],
    compareToken: 0,
    spacePan: false,
    toastTimer: null
  };

  function currentItem() {
    return state.items[state.current] || null;
  }

  function cloneMarkers(markers) {
    return markers.map((marker) => ({ ...marker }));
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function storageKey(item) {
    return STORAGE_PREFIX + encodeURIComponent(item.id);
  }

  function readSavedMarkers(item) {
    try {
      const value = JSON.parse(localStorage.getItem(storageKey(item)) || "[]");
      return Array.isArray(value) ? value.filter(validMarker) : [];
    } catch {
      return [];
    }
  }

  function validMarker(marker) {
    if (!marker || !["arrow", "circle"].includes(marker.type) || typeof marker.id !== "number") return false;
    const numbers = marker.type === "circle"
      ? [marker.cx, marker.cy, marker.r, marker.lineWidth]
      : [marker.x1, marker.y1, marker.x2, marker.y2, marker.lineWidth];
    return numbers.every(Number.isFinite) && typeof marker.color === "string";
  }

  function persist(item) {
    try {
      localStorage.setItem(storageKey(item), JSON.stringify(item.markers));
    } catch {
      showToast("浏览器无法自动保存，但仍可继续标注和下载", true);
    }
  }

  function showToast(message, error = false) {
    clearTimeout(state.toastTimer);
    el.toast.textContent = message;
    el.toast.classList.toggle("error", error);
    el.toast.classList.add("show");
    state.toastTimer = setTimeout(() => el.toast.classList.remove("show"), 2600);
  }

  function fileId(file) {
    return `${file.webkitRelativePath || file.name}|${file.size}|${file.lastModified}`;
  }

  function baseName(name) {
    return name.replace(/\.[^.]+$/, "");
  }

  function clipboardFile(blob, index = 0) {
    const now = new Date();
    const stamp = [now.getFullYear(), now.getMonth() + 1, now.getDate(), now.getHours(), now.getMinutes(), now.getSeconds()]
      .map((part, position) => position ? String(part).padStart(2, "0") : part)
      .join("");
    const suffix = index ? `_${index + 1}` : "";
    return new File([blob], `PPT粘贴_${stamp}${suffix}.png`, {
      type: blob.type || "image/png",
      lastModified: Date.now() + index
    });
  }

  function addClipboardImages(blobs) {
    if (!blobs.length) {
      showToast("剪贴板中没有图片，请先在 PowerPoint 中复制图片", true);
      return;
    }
    addFiles(blobs.map(clipboardFile));
    showToast(`已从剪贴板粘贴 ${blobs.length} 张图片`);
  }

  async function pasteFromClipboard() {
    if (!navigator.clipboard?.read) {
      showToast("请在 PowerPoint 复制图片后，直接按 Ctrl/⌘+V", true);
      return;
    }
    try {
      const clipboardItems = await navigator.clipboard.read();
      const blobs = [];
      for (const item of clipboardItems) {
        const imageType = item.types.find((type) => type.startsWith("image/"));
        if (imageType) blobs.push(await item.getType(imageType));
      }
      addClipboardImages(blobs);
    } catch {
      showToast("浏览器未允许读取剪贴板，请直接按 Ctrl/⌘+V", true);
    }
  }

  function addFiles(fileList) {
    const existing = new Set(state.items.map((item) => item.id));
    const additions = [...fileList]
      .filter((file) => SUPPORTED.test(file.name) && !existing.has(fileId(file)))
      .sort((a, b) => a.name.localeCompare(b.name, "zh-CN", { numeric: true }));

    for (const file of additions) {
      const item = {
        id: fileId(file),
        file,
        name: file.name,
        markers: [],
        undo: [],
        redo: [],
        width: 0,
        height: 0,
        template: { arrowOffset: null, circleRadius: null }
      };
      item.markers = readSavedMarkers(item);
      state.items.push(item);
    }

    if (!additions.length) {
      if (fileList.length) showToast("没有发现新的受支持图片", true);
      return;
    }

    renderImageList();
    if (state.current === -1) openItem(0);
    else {
      renderComparePanels();
      showToast(`已加入 ${additions.length} 张图片`);
    }
  }

  async function openItem(index) {
    if (index < 0 || index >= state.items.length || index === state.current && state.loadedId === currentItem()?.id) return;

    state.current = index;
    state.selectedId = null;
    state.interaction = null;
    renderImageList();
    updateControls();

    const item = currentItem();
    const token = ++state.loadToken;
    el.empty.hidden = true;
    el.loading.hidden = false;
    el.stage.hidden = true;
    el.currentName.textContent = item.name;
    el.currentMeta.textContent = "正在读取图片…";

    try {
      const decodedCanvas = document.createElement("canvas");
      const dimensions = await decodeToCanvas(item.file, decodedCanvas);
      if (token !== state.loadToken) {
        decodedCanvas.width = decodedCanvas.height = 1;
        return;
      }
      item.width = dimensions.width;
      item.height = dimensions.height;
      state.loadedId = item.id;

      el.canvas.width = item.width;
      el.canvas.height = item.height;
      el.canvas.getContext("2d", { alpha: false }).drawImage(decodedCanvas, 0, 0);
      decodedCanvas.width = decodedCanvas.height = 1;

      el.stage.style.width = `${item.width}px`;
      el.stage.style.height = `${item.height}px`;
      el.overlay.setAttribute("viewBox", `0 0 ${item.width} ${item.height}`);
      el.overlay.setAttribute("width", item.width);
      el.overlay.setAttribute("height", item.height);
      el.loading.hidden = true;
      el.stage.hidden = false;
      el.currentMeta.textContent = `${item.width} × ${item.height} px · ${item.markers.length} 个标记`;
      requestAnimationFrame(fitToView);
      renderMarkers();
      updateControls();
      renderComparePanels();
    } catch (error) {
      if (token !== state.loadToken) return;
      state.loadedId = null;
      el.loading.hidden = true;
      el.stage.hidden = true;
      el.empty.hidden = false;
      el.currentMeta.textContent = "读取失败";
      showToast(error?.message || "无法读取这张图片", true);
      console.error(error);
    }
  }

  async function decodeToCanvas(file, canvas) {
    const context = canvas.getContext("2d", { alpha: false });
    if (/\.tiff?$/i.test(file.name)) {
      if (!window.UTIF) throw new Error("TIFF 解码器未加载，请刷新页面重试");
      const buffer = await file.arrayBuffer();
      const pages = window.UTIF.decode(buffer);
      if (!pages.length) throw new Error("这个 TIFF 文件中没有可读取的图像");
      window.UTIF.decodeImage(buffer, pages[0]);
      const rgba = window.UTIF.toRGBA8(pages[0]);
      const width = pages[0].width;
      const height = pages[0].height;
      if (!width || !height || width * height > 120_000_000) throw new Error("图片尺寸过大，浏览器无法安全打开");
      canvas.width = width;
      canvas.height = height;
      const pixels = new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, rgba.byteLength);
      context.putImageData(new ImageData(pixels, width, height), 0, 0);
      return { width, height };
    }

    const bitmap = await createImageBitmap(file, { imageOrientation: "from-image" });
    if (bitmap.width * bitmap.height > 120_000_000) {
      bitmap.close();
      throw new Error("图片尺寸过大，浏览器无法安全打开");
    }
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    context.drawImage(bitmap, 0, 0);
    const dimensions = { width: bitmap.width, height: bitmap.height };
    bitmap.close();
    return dimensions;
  }

  function nextCompareIndex(slot) {
    const used = new Set([state.current, ...state.compareIndices.slice(0, slot)]);
    for (let index = 0; index < state.items.length; index++) {
      if (!used.has(index)) return index;
    }
    return -1;
  }

  function ensureCompareIndices() {
    for (let slot = 0; slot < state.splitCount - 1; slot++) {
      const index = state.compareIndices[slot];
      if (index < 0 || index >= state.items.length) state.compareIndices[slot] = nextCompareIndex(slot);
    }
  }

  function setSplitCount(value) {
    state.splitCount = clamp(Number(value) || 1, 1, 4);
    el.splitCountInput.value = String(state.splitCount);
    el.splitViewport.classList.remove("split-1", "split-2", "split-3", "split-4");
    el.splitViewport.classList.add(`split-${state.splitCount}`);
    el.primaryBadge.hidden = state.splitCount === 1;
    ensureCompareIndices();
    renderComparePanels();
    requestAnimationFrame(fitToView);
  }

  function renderComparePanels() {
    const token = ++state.compareToken;
    el.splitViewport.querySelectorAll(".compare-panel").forEach((panel) => panel.remove());
    if (state.splitCount === 1) return;
    ensureCompareIndices();

    const panels = [];
    for (let slot = 0; slot < state.splitCount - 1; slot++) {
      const panel = document.createElement("section");
      panel.className = "compare-panel";
      panel.dataset.slot = slot;
      panel.setAttribute("aria-label", `对照屏 ${slot + 2}`);

      const bar = document.createElement("div");
      bar.className = "compare-panel-bar";
      const label = document.createElement("span");
      label.className = "screen-label";
      label.textContent = `屏 ${slot + 2}`;

      const select = document.createElement("select");
      select.className = "compare-select";
      select.setAttribute("aria-label", `选择对照屏 ${slot + 2} 的图片`);
      const placeholder = document.createElement("option");
      placeholder.value = "";
      placeholder.textContent = "选择对照图片";
      select.append(placeholder);
      state.items.forEach((item, index) => {
        const option = document.createElement("option");
        option.value = index;
        option.textContent = `${String(index + 1).padStart(2, "0")} · ${item.name}`;
        select.append(option);
      });
      const selectedIndex = state.compareIndices[slot];
      select.value = selectedIndex >= 0 ? String(selectedIndex) : "";
      select.addEventListener("change", () => {
        state.compareIndices[slot] = select.value === "" ? -1 : Number(select.value);
        renderComparePanels();
      });

      const editButton = document.createElement("button");
      editButton.type = "button";
      editButton.className = "edit-compare-button";
      editButton.textContent = "编辑此图";
      editButton.disabled = selectedIndex < 0 || selectedIndex === state.current;
      editButton.addEventListener("click", () => editComparedImage(slot));
      bar.append(label, select, editButton);

      const wrap = document.createElement("div");
      wrap.className = "compare-canvas-wrap";
      const canvas = document.createElement("canvas");
      canvas.className = "compare-canvas";
      canvas.hidden = true;
      const status = document.createElement("div");
      status.className = "compare-status";
      status.textContent = selectedIndex < 0 ? "请先添加并选择一张对照图片" : "正在读取对照图片…";
      wrap.append(canvas, status);
      panel.append(bar, wrap);
      el.splitViewport.append(panel);
      panels.push({ panel, slot });
    }

    (async () => {
      for (const entry of panels) {
        if (token !== state.compareToken) return;
        await renderCompareCanvas(entry.panel, entry.slot, token);
      }
    })();
  }

  async function renderCompareCanvas(panel, slot, token = state.compareToken) {
    const index = state.compareIndices[slot];
    const canvas = panel.querySelector("canvas");
    const status = panel.querySelector(".compare-status");
    if (index < 0 || !state.items[index]) return;
    const item = state.items[index];

    try {
      let source = el.canvas;
      let dimensions = { width: item.width, height: item.height };
      if (index !== state.current || state.loadedId !== item.id) {
        source = document.createElement("canvas");
        dimensions = await decodeToCanvas(item.file, source);
      }
      if (token !== state.compareToken || !panel.isConnected || state.compareIndices[slot] !== index) return;
      item.width = dimensions.width;
      item.height = dimensions.height;
      canvas.width = item.width;
      canvas.height = item.height;
      const context = canvas.getContext("2d", { alpha: false });
      context.drawImage(source, 0, 0);
      item.markers.forEach((marker) => drawMarker(context, marker));
      canvas.hidden = false;
      status.hidden = true;
    } catch (error) {
      if (token !== state.compareToken || !panel.isConnected) return;
      status.textContent = "无法读取这张对照图片";
      console.error(error);
    }
  }

  async function editComparedImage(slot) {
    const target = state.compareIndices[slot];
    if (target < 0 || target === state.current || !state.items[target]) return;
    const previous = state.current;
    state.compareIndices[slot] = previous;
    await openItem(target);
  }

  function refreshCurrentComparisons() {
    el.splitViewport.querySelectorAll(".compare-panel").forEach((panel) => {
      const slot = Number(panel.dataset.slot);
      if (state.compareIndices[slot] === state.current) renderCompareCanvas(panel, slot);
    });
  }

  function renderImageList() {
    el.imageCount.textContent = state.items.length;
    el.imageList.replaceChildren();
    state.items.forEach((item, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `image-item${index === state.current ? " active" : ""}`;
      button.setAttribute("role", "listitem");
      button.setAttribute("aria-label", `打开 ${item.name}，${item.markers.length} 个标记`);
      button.innerHTML = `
        <span class="image-index">${String(index + 1).padStart(2, "0")}</span>
        <span class="image-copy"><strong></strong><small>${formatBytes(item.file.size)}</small></span>
        <span class="marker-count" title="标记数量">${item.markers.length}</span>`;
      button.querySelector("strong").textContent = item.name;
      button.addEventListener("click", () => openItem(index));
      el.imageList.append(button);
    });
  }

  function formatBytes(bytes) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function fitToView() {
    const item = currentItem();
    if (!item?.width || state.loadedId !== item.id) return;
    const bounds = el.viewport.getBoundingClientRect();
    state.fitScale = Math.min(1, (bounds.width - 52) / item.width, (bounds.height - 52) / item.height);
    state.scale = Math.max(0.02, state.fitScale);
    state.panX = (bounds.width - item.width * state.scale) / 2;
    state.panY = (bounds.height - item.height * state.scale) / 2;
    applyTransform();
  }

  function applyTransform() {
    el.stage.style.transform = `translate(${state.panX}px, ${state.panY}px) scale(${state.scale})`;
    el.zoomOutput.value = `${Math.round(state.scale * 100)}%`;
    el.zoomOutput.textContent = `${Math.round(state.scale * 100)}%`;
    renderMarkers();
  }

  function zoomAt(factor, clientX, clientY) {
    const item = currentItem();
    if (!item?.width || state.loadedId !== item.id) return;
    const rect = el.viewport.getBoundingClientRect();
    const x = (clientX ?? rect.left + rect.width / 2) - rect.left;
    const y = (clientY ?? rect.top + rect.height / 2) - rect.top;
    const imageX = (x - state.panX) / state.scale;
    const imageY = (y - state.panY) / state.scale;
    const next = clamp(state.scale * factor, Math.max(0.01, state.fitScale * 0.25), 8);
    state.panX = x - imageX * next;
    state.panY = y - imageY * next;
    state.scale = next;
    applyTransform();
  }

  function clientToImage(clientX, clientY) {
    const rect = el.viewport.getBoundingClientRect();
    return {
      x: (clientX - rect.left - state.panX) / state.scale,
      y: (clientY - rect.top - state.panY) / state.scale
    };
  }

  function insideImage(point, item = currentItem()) {
    return item && point.x >= 0 && point.y >= 0 && point.x <= item.width && point.y <= item.height;
  }

  function setTool(tool) {
    state.tool = tool;
    if (tool !== "select") state.selectedId = null;
    $$(".tool-button").forEach((button) => button.classList.toggle("active", button.dataset.tool === tool));
    el.viewport.className = `viewport tool-${tool}${state.spacePan ? " space-pan" : ""}`;
    const hints = {
      arrow: "箭头：单击快速放置 · 沿箭头方向拖动",
      circle: "圆圈：单击快速放置 · 拖动可指定大小",
      select: "选择：拖动标记移动 · 拖动控制点调整",
      pan: "移动：拖动画布 · 滚轮缩放"
    };
    el.hint.textContent = state.pathology && tool === "circle"
      ? "CIC 手动圈选：滚轮放大 · 从中心拖出圆圈，随后单击重复 · 按住空格拖动视野"
      : hints[tool];
    renderMarkers();
    updateControls();
  }

  function defaultRadius(item) {
    if (state.pathology) return (8 + state.size * 2) / state.scale;
    return Math.min(item.width, item.height) * (0.012 + state.size * 0.0032);
  }

  function defaultArrowOffset(item) {
    const length = defaultRadius(item) * 2.7;
    return { x: -length / Math.SQRT2, y: -length / Math.SQRT2 };
  }

  function currentLineWidth(item) {
    if (state.pathology) return 0.5 + state.stroke * 0.5;
    return Math.min(item.width, item.height) * (0.0011 + state.stroke * 0.00043);
  }

  function nextMarkerId(item) {
    return Math.max(0, ...item.markers.map((marker) => marker.id)) + 1;
  }

  function createMarker(type, point, item) {
    const common = { id: nextMarkerId(item), type, color: state.color, lineWidth: currentLineWidth(item) };
    if (type === "circle") {
      return { ...common, cic: state.pathology, cx: point.x, cy: point.y, r: item.template.circleRadius || defaultRadius(item) };
    }
    const offset = item.template.arrowOffset || defaultArrowOffset(item);
    let x1 = point.x + offset.x;
    let y1 = point.y + offset.y;
    if (x1 < 0 || x1 > item.width) x1 = point.x - offset.x;
    if (y1 < 0 || y1 > item.height) y1 = point.y - offset.y;
    return {
      ...common,
      x1: clamp(x1, 0, item.width),
      y1: clamp(y1, 0, item.height),
      x2: point.x,
      y2: point.y
    };
  }

  function arrowGeometry(marker) {
    const dx = marker.x2 - marker.x1;
    const dy = marker.y2 - marker.y1;
    const length = Math.max(0.001, Math.hypot(dx, dy));
    const ux = dx / length;
    const uy = dy / length;
    const headLength = Math.min(length * 0.44, Math.max(marker.lineWidth * 5, 12));
    const headWidth = Math.max(marker.lineWidth * 3.2, headLength * 0.68);
    const baseX = marker.x2 - ux * headLength;
    const baseY = marker.y2 - uy * headLength;
    const px = -uy * headWidth / 2;
    const py = ux * headWidth / 2;
    return {
      baseX,
      baseY,
      leftX: baseX + px,
      leftY: baseY + py,
      rightX: baseX - px,
      rightY: baseY - py
    };
  }

  function svgElement(name, attributes = {}) {
    const node = document.createElementNS(SVG_NS, name);
    Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, value));
    return node;
  }

  function renderMarkers() {
    const item = currentItem();
    el.overlay.replaceChildren();
    if (!item?.width || state.loadedId !== item.id) return;

    for (const marker of item.markers) {
      const outline = Math.max(2, marker.lineWidth * 0.5);
      if (marker.type === "circle") {
        if (!marker.cic) el.overlay.append(
          svgElement("circle", { cx: marker.cx, cy: marker.cy, r: marker.r, fill: "none", stroke: "#0a0d0e", "stroke-width": marker.lineWidth + outline * 2 })
        );
        el.overlay.append(
          svgElement("circle", { cx: marker.cx, cy: marker.cy, r: marker.r, fill: "none", stroke: marker.color, "stroke-width": marker.lineWidth })
        );
      } else {
        const geometry = arrowGeometry(marker);
        el.overlay.append(
          svgElement("line", { x1: marker.x1, y1: marker.y1, x2: geometry.baseX, y2: geometry.baseY, stroke: "#0a0d0e", "stroke-width": marker.lineWidth + outline * 2, "stroke-linecap": "round" }),
          svgElement("line", { x1: marker.x1, y1: marker.y1, x2: geometry.baseX, y2: geometry.baseY, stroke: marker.color, "stroke-width": marker.lineWidth, "stroke-linecap": "round" }),
          svgElement("polygon", {
            points: `${marker.x2},${marker.y2} ${geometry.leftX},${geometry.leftY} ${geometry.rightX},${geometry.rightY}`,
            fill: marker.color,
            stroke: "#0a0d0e",
            "stroke-width": outline * 1.4,
            "stroke-linejoin": "round"
          })
        );
      }
    }

    if (state.tool === "select" && state.selectedId !== null) renderSelection(item.markers.find((marker) => marker.id === state.selectedId));
  }

  function renderSelection(marker) {
    if (!marker) return;
    const unit = 1 / state.scale;
    const handleRadius = 6.5 * unit;
    const attrs = { fill: "#0d1717", stroke: "#55e6ff", "stroke-width": 2 * unit };

    if (marker.type === "circle") {
      el.overlay.append(
        svgElement("circle", {
          cx: marker.cx,
          cy: marker.cy,
          r: marker.r + 5 * unit,
          fill: "none",
          stroke: "#55e6ff",
          "stroke-width": 1.5 * unit,
          "stroke-dasharray": `${6 * unit} ${5 * unit}`
        }),
        svgElement("circle", { cx: marker.cx + marker.r, cy: marker.cy, r: handleRadius, ...attrs })
      );
    } else {
      el.overlay.append(
        svgElement("line", {
          x1: marker.x1,
          y1: marker.y1,
          x2: marker.x2,
          y2: marker.y2,
          stroke: "#55e6ff",
          "stroke-width": 1.5 * unit,
          "stroke-dasharray": `${6 * unit} ${5 * unit}`
        }),
        svgElement("circle", { cx: marker.x1, cy: marker.y1, r: handleRadius, ...attrs }),
        svgElement("circle", { cx: marker.x2, cy: marker.y2, r: handleRadius, ...attrs })
      );
    }
  }

  function distanceToSegment(point, a, b) {
    const dx = b.x - a.x;
    const dy = b.y - a.y;
    const lengthSquared = dx * dx + dy * dy;
    if (!lengthSquared) return Math.hypot(point.x - a.x, point.y - a.y);
    const t = clamp(((point.x - a.x) * dx + (point.y - a.y) * dy) / lengthSquared, 0, 1);
    return Math.hypot(point.x - (a.x + t * dx), point.y - (a.y + t * dy));
  }

  function hitHandle(point, marker) {
    if (!marker) return null;
    const tolerance = 12 / state.scale;
    if (marker.type === "circle") {
      return Math.hypot(point.x - (marker.cx + marker.r), point.y - marker.cy) <= tolerance ? "radius" : null;
    }
    if (Math.hypot(point.x - marker.x1, point.y - marker.y1) <= tolerance) return "tail";
    if (Math.hypot(point.x - marker.x2, point.y - marker.y2) <= tolerance) return "tip";
    return null;
  }

  function hitMarker(point, item) {
    const tolerance = 11 / state.scale;
    for (let index = item.markers.length - 1; index >= 0; index--) {
      const marker = item.markers[index];
      if (marker.type === "circle") {
        if (Math.hypot(point.x - marker.cx, point.y - marker.cy) <= marker.r + tolerance) return marker;
      } else if (distanceToSegment(point, { x: marker.x1, y: marker.y1 }, { x: marker.x2, y: marker.y2 }) <= tolerance + marker.lineWidth / 2) {
        return marker;
      }
    }
    return null;
  }

  function pointerDown(event) {
    const item = currentItem();
    if (!item?.width || state.loadedId !== item.id || event.button > 1) return;
    el.viewport.focus({ preventScroll: true });

    if (event.button === 1 || state.spacePan || state.tool === "pan") {
      state.interaction = { mode: "pan", clientX: event.clientX, clientY: event.clientY, panX: state.panX, panY: state.panY };
      el.viewport.classList.add("panning");
      el.viewport.setPointerCapture(event.pointerId);
      event.preventDefault();
      return;
    }

    const point = clientToImage(event.clientX, event.clientY);
    if (!insideImage(point, item)) return;

    if (state.tool === "arrow" || state.tool === "circle") {
      const before = cloneMarkers(item.markers);
      const marker = createMarker(state.tool, point, item);
      item.markers.push(marker);
      state.selectedId = marker.id;
      state.interaction = { mode: "create", markerId: marker.id, type: marker.type, start: point, startClientX: event.clientX, startClientY: event.clientY, before, moved: false };
      el.viewport.setPointerCapture(event.pointerId);
      renderMarkers();
      updateControls();
      event.preventDefault();
      return;
    }

    if (state.tool === "select") {
      const selected = item.markers.find((marker) => marker.id === state.selectedId);
      const handle = hitHandle(point, selected);
      const marker = selected && handle ? selected : hitMarker(point, item);
      if (!marker) {
        state.selectedId = null;
        renderMarkers();
        updateControls();
        return;
      }
      state.selectedId = marker.id;
      state.interaction = {
        mode: handle || "move",
        markerId: marker.id,
        start: point,
        original: { ...marker },
        before: cloneMarkers(item.markers),
        changed: false
      };
      el.viewport.setPointerCapture(event.pointerId);
      renderMarkers();
      updateControls();
      event.preventDefault();
    }
  }

  function pointerMove(event) {
    const interaction = state.interaction;
    const item = currentItem();
    if (!interaction || !item) return;

    if (interaction.mode === "pan") {
      state.panX = interaction.panX + event.clientX - interaction.clientX;
      state.panY = interaction.panY + event.clientY - interaction.clientY;
      applyTransform();
      return;
    }

    const marker = item.markers.find((candidate) => candidate.id === interaction.markerId);
    if (!marker) return;
    const point = clientToImage(event.clientX, event.clientY);

    if (interaction.mode === "create") {
      const moved = Math.hypot(event.clientX - interaction.startClientX, event.clientY - interaction.startClientY) > 7;
      if (!moved) return;
      interaction.moved = true;
      if (marker.type === "circle") {
        marker.r = clamp(Math.hypot(point.x - marker.cx, point.y - marker.cy), 4 / state.scale, Math.hypot(item.width, item.height));
      } else {
        marker.x1 = interaction.start.x;
        marker.y1 = interaction.start.y;
        marker.x2 = clamp(point.x, 0, item.width);
        marker.y2 = clamp(point.y, 0, item.height);
      }
    } else if (interaction.mode === "move") {
      const dx = point.x - interaction.start.x;
      const dy = point.y - interaction.start.y;
      if (marker.type === "circle") {
        marker.cx = clamp(interaction.original.cx + dx, 0, item.width);
        marker.cy = clamp(interaction.original.cy + dy, 0, item.height);
      } else {
        const limitedX = clamp(dx, -Math.min(interaction.original.x1, interaction.original.x2), item.width - Math.max(interaction.original.x1, interaction.original.x2));
        const limitedY = clamp(dy, -Math.min(interaction.original.y1, interaction.original.y2), item.height - Math.max(interaction.original.y1, interaction.original.y2));
        marker.x1 = interaction.original.x1 + limitedX;
        marker.y1 = interaction.original.y1 + limitedY;
        marker.x2 = interaction.original.x2 + limitedX;
        marker.y2 = interaction.original.y2 + limitedY;
      }
      interaction.changed = Math.abs(dx) + Math.abs(dy) > 0.01;
    } else if (interaction.mode === "radius") {
      marker.r = clamp(Math.hypot(point.x - marker.cx, point.y - marker.cy), 4 / state.scale, Math.hypot(item.width, item.height));
      interaction.changed = true;
    } else if (interaction.mode === "tail") {
      marker.x1 = clamp(point.x, 0, item.width);
      marker.y1 = clamp(point.y, 0, item.height);
      interaction.changed = true;
    } else if (interaction.mode === "tip") {
      marker.x2 = clamp(point.x, 0, item.width);
      marker.y2 = clamp(point.y, 0, item.height);
      interaction.changed = true;
    }
    renderMarkers();
  }

  function pointerUp(event) {
    const interaction = state.interaction;
    const item = currentItem();
    if (!interaction || !item) return;
    if (el.viewport.hasPointerCapture(event.pointerId)) el.viewport.releasePointerCapture(event.pointerId);
    el.viewport.classList.remove("panning");

    if (interaction.mode === "create") {
      pushUndo(item, interaction.before);
      const marker = item.markers.find((candidate) => candidate.id === interaction.markerId);
      rememberTemplate(item, marker);
      persist(item);
      showToast(marker?.type === "arrow" ? "已添加箭头" : "已添加圆圈");
    } else if (interaction.mode !== "pan" && interaction.changed) {
      pushUndo(item, interaction.before);
      rememberTemplate(item, item.markers.find((marker) => marker.id === interaction.markerId));
      persist(item);
    }

    state.interaction = null;
    renderImageList();
    renderMarkers();
    updateControls();
    refreshCurrentComparisons();
  }

  function pointerCancel(event) {
    const item = currentItem();
    if (!state.interaction || !item) return;
    if (state.interaction.before) item.markers = state.interaction.before;
    if (el.viewport.hasPointerCapture(event.pointerId)) el.viewport.releasePointerCapture(event.pointerId);
    el.viewport.classList.remove("panning");
    state.interaction = null;
    renderMarkers();
    updateControls();
  }

  function rememberTemplate(item, marker) {
    if (!marker) return;
    if (marker.type === "circle") item.template.circleRadius = marker.r;
    else item.template.arrowOffset = { x: marker.x1 - marker.x2, y: marker.y1 - marker.y2 };
  }

  function pushUndo(item, snapshot) {
    item.undo.push(snapshot);
    if (item.undo.length > HISTORY_LIMIT) item.undo.shift();
    item.redo = [];
  }

  function undo() {
    const item = currentItem();
    if (!item?.undo.length) return;
    item.redo.push(cloneMarkers(item.markers));
    item.markers = item.undo.pop();
    state.selectedId = null;
    persist(item);
    refreshAfterEdit();
  }

  function redo() {
    const item = currentItem();
    if (!item?.redo.length) return;
    item.undo.push(cloneMarkers(item.markers));
    item.markers = item.redo.pop();
    state.selectedId = null;
    persist(item);
    refreshAfterEdit();
  }

  function deleteSelected() {
    const item = currentItem();
    if (!item || state.selectedId === null) return;
    const index = item.markers.findIndex((marker) => marker.id === state.selectedId);
    if (index < 0) return;
    pushUndo(item, cloneMarkers(item.markers));
    item.markers.splice(index, 1);
    state.selectedId = null;
    persist(item);
    refreshAfterEdit();
    showToast("已删除标记");
  }

  function clearMarkers() {
    const item = currentItem();
    if (!item?.markers.length || !confirm(`清空“${item.name}”上的全部 ${item.markers.length} 个标记？`)) return;
    pushUndo(item, cloneMarkers(item.markers));
    item.markers = [];
    state.selectedId = null;
    persist(item);
    refreshAfterEdit();
    showToast("已清空当前图片的标记");
  }

  function refreshAfterEdit() {
    renderImageList();
    renderMarkers();
    updateControls();
    refreshCurrentComparisons();
  }

  function updateSelectedStyle(change) {
    const item = currentItem();
    if (state.tool !== "select" || state.selectedId === null || !item) return;
    const marker = item.markers.find((candidate) => candidate.id === state.selectedId);
    if (!marker) return;
    const before = cloneMarkers(item.markers);
    change(marker, item);
    pushUndo(item, before);
    rememberTemplate(item, marker);
    persist(item);
    refreshAfterEdit();
  }

  function updateControls() {
    const item = currentItem();
    const ready = Boolean(item?.width && state.loadedId === item.id);
    el.undo.disabled = !item?.undo.length;
    el.redo.disabled = !item?.redo.length;
    el.delete.disabled = !(ready && state.tool === "select" && state.selectedId !== null);
    el.clear.disabled = !item?.markers.length;
    el.copy.disabled = !ready;
    el.export.disabled = !ready;
    el.exportNext.disabled = !ready;
    if (ready) {
      const cicCount = item.markers.filter((marker) => marker.cic === true).length;
      el.currentMeta.textContent = `${item.width} × ${item.height} px · ${item.markers.length} 个标记`
        + (state.pathology || cicCount ? ` · CIC 手动圈选 ${cicCount}` : "");
    }
  }

  function drawMarker(context, marker) {
    const outline = Math.max(2, marker.lineWidth * 0.5);
    context.save();
    context.lineCap = "round";
    context.lineJoin = "round";
    if (marker.type === "circle") {
      context.beginPath();
      context.arc(marker.cx, marker.cy, marker.r, 0, Math.PI * 2);
      if (!marker.cic) {
        context.strokeStyle = "#0a0d0e";
        context.lineWidth = marker.lineWidth + outline * 2;
        context.stroke();
      }
      context.strokeStyle = marker.color;
      context.lineWidth = marker.lineWidth;
      context.stroke();
    } else {
      const geometry = arrowGeometry(marker);
      context.beginPath();
      context.moveTo(marker.x1, marker.y1);
      context.lineTo(geometry.baseX, geometry.baseY);
      context.strokeStyle = "#0a0d0e";
      context.lineWidth = marker.lineWidth + outline * 2;
      context.stroke();
      context.strokeStyle = marker.color;
      context.lineWidth = marker.lineWidth;
      context.stroke();

      context.beginPath();
      context.moveTo(marker.x2, marker.y2);
      context.lineTo(geometry.leftX, geometry.leftY);
      context.lineTo(geometry.rightX, geometry.rightY);
      context.closePath();
      context.fillStyle = marker.color;
      context.strokeStyle = "#0a0d0e";
      context.lineWidth = outline * 1.4;
      context.stroke();
      context.fill();
    }
    context.restore();
  }

  function annotatedCanvas() {
    const item = currentItem();
    if (!item?.width || state.loadedId !== item.id) throw new Error("请先选择一张图片");
    const output = document.createElement("canvas");
    output.width = item.width;
    output.height = item.height;
    const context = output.getContext("2d", { alpha: false });
    context.drawImage(el.canvas, 0, 0);
    item.markers.forEach((marker) => drawMarker(context, marker));
    return output;
  }

  function pngBlob() {
    return new Promise((resolve, reject) => {
      const output = annotatedCanvas();
      output.toBlob((value) => {
        output.width = output.height = 1;
        value ? resolve(value) : reject(new Error("PNG 生成失败"));
      }, "image/png");
    });
  }

  async function copyAnnotatedImage() {
    if (!navigator.clipboard?.write || !window.ClipboardItem) {
      showToast("当前浏览器不支持复制图片，请使用下载 PNG", true);
      return;
    }
    el.copy.disabled = true;
    try {
      await navigator.clipboard.write([new ClipboardItem({ "image/png": pngBlob() })]);
      showToast("标注图已复制，回到 PowerPoint 按 Ctrl/⌘+V 即可粘贴");
    } catch {
      showToast("浏览器未允许复制图片，请重试或使用下载 PNG", true);
    } finally {
      updateControls();
    }
  }

  async function exportPng(advance = false) {
    const item = currentItem();
    if (!item?.width || state.loadedId !== item.id) return;
    el.export.disabled = true;
    el.exportNext.disabled = true;

    try {
      const blob = await pngBlob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `${baseName(item.name)}_${item.markers.some((marker) => marker.cic) || state.pathology ? "CIC标注" : "融合标注"}.png`;
      link.click();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      showToast(`已下载 ${link.download}`);
      if (advance && state.current < state.items.length - 1) await openItem(state.current + 1);
    } catch (error) {
      showToast(error?.message || "下载失败", true);
    } finally {
      updateControls();
    }
  }

  function handleKeyDown(event) {
    const tag = event.target.tagName;
    if (["INPUT", "TEXTAREA", "SELECT"].includes(tag)) return;
    const modifier = event.ctrlKey || event.metaKey;
    if (modifier && event.key.toLowerCase() === "z") {
      event.preventDefault();
      event.shiftKey ? redo() : undo();
      return;
    }
    if (modifier && event.key.toLowerCase() === "y") {
      event.preventDefault();
      redo();
      return;
    }
    if (modifier) return;
    if (event.code === "Space" && !event.repeat) {
      event.preventDefault();
      state.spacePan = true;
      el.viewport.classList.add("space-pan");
      return;
    }
    const key = event.key.toLowerCase();
    if ({ v: "select", a: "arrow", c: "circle", h: "pan" }[key]) setTool({ v: "select", a: "arrow", c: "circle", h: "pan" }[key]);
    else if (key === "delete" || key === "backspace") { event.preventDefault(); deleteSelected(); }
    else if (key === "escape") { state.selectedId = null; renderMarkers(); updateControls(); }
    else if (key === "0") fitToView();
    else if (key === "+" || key === "=") zoomAt(1.2);
    else if (key === "-") zoomAt(1 / 1.2);
    else if (key === "e" && currentItem()) exportPng(false);
  }

  function runSelfCheck() {
    if (!new URLSearchParams(location.search).has("selftest")) return;
    const geometry = arrowGeometry({ x1: 0, y1: 0, x2: 100, y2: 0, lineWidth: 5 });
    const ok = Math.abs(distanceToSegment({ x: 50, y: 10 }, { x: 0, y: 0 }, { x: 100, y: 0 }) - 10) < 0.001
      && geometry.baseX < 100 && geometry.leftY !== geometry.rightY;
    document.documentElement.dataset.selftest = ok ? "passed" : "failed";
    if (!ok) throw new Error("FusionMark geometry self-check failed");
  }

  el.fileInput.addEventListener("change", (event) => {
    addFiles(event.target.files);
    event.target.value = "";
  });

  el.pasteButton.addEventListener("click", pasteFromClipboard);
  el.annotationMode.addEventListener("change", () => {
    state.pathology = el.annotationMode.value === "cic";
    state.color = state.pathology ? "#d00000" : "#ffffff";
    state.items.forEach((item) => { item.template.circleRadius = null; });
    $$(".swatch").forEach((button) => button.classList.toggle("active", button.dataset.color === state.color));
    setTool(state.pathology ? "circle" : "arrow");
    showToast(state.pathology ? "病理 CIC：请放大后手动圈选，拖动一次可记住圆圈大小" : "已切换到荧光标注");
  });
  window.addEventListener("paste", (event) => {
    const blobs = [...(event.clipboardData?.items || [])]
      .filter((item) => item.kind === "file" && item.type.startsWith("image/"))
      .map((item) => item.getAsFile())
      .filter(Boolean);
    if (!blobs.length) return;
    event.preventDefault();
    addClipboardImages(blobs);
  });

  [el.dropZone, el.splitViewport].forEach((target) => {
    target.addEventListener("dragover", (event) => { event.preventDefault(); el.dropZone.classList.add("dragging"); });
    target.addEventListener("dragleave", () => el.dropZone.classList.remove("dragging"));
    target.addEventListener("drop", (event) => {
      event.preventDefault();
      el.dropZone.classList.remove("dragging");
      addFiles(event.dataTransfer.files);
    });
  });

  $$(".tool-button").forEach((button) => button.addEventListener("click", () => setTool(button.dataset.tool)));
  $$(".swatch").forEach((button) => button.addEventListener("click", () => {
    state.color = button.dataset.color;
    $$(".swatch").forEach((swatch) => swatch.classList.toggle("active", swatch === button));
    updateSelectedStyle((marker) => { marker.color = state.color; });
  }));

  el.sizeInput.addEventListener("input", () => { el.sizeOutput.value = el.sizeInput.value; el.sizeOutput.textContent = el.sizeInput.value; });
  el.sizeInput.addEventListener("change", () => {
    state.size = Number(el.sizeInput.value);
    const item = currentItem();
    if (item) item.template = { arrowOffset: null, circleRadius: null };
    updateSelectedStyle((marker, selectedItem) => {
      if (marker.type === "circle") marker.r = defaultRadius(selectedItem);
      else {
        const length = defaultRadius(selectedItem) * 2.7;
        const currentLength = Math.max(0.001, Math.hypot(marker.x2 - marker.x1, marker.y2 - marker.y1));
        marker.x1 = clamp(marker.x2 - (marker.x2 - marker.x1) / currentLength * length, 0, selectedItem.width);
        marker.y1 = clamp(marker.y2 - (marker.y2 - marker.y1) / currentLength * length, 0, selectedItem.height);
      }
    });
  });

  el.strokeInput.addEventListener("input", () => { el.strokeOutput.value = el.strokeInput.value; el.strokeOutput.textContent = el.strokeInput.value; });
  el.strokeInput.addEventListener("change", () => {
    state.stroke = Number(el.strokeInput.value);
    updateSelectedStyle((marker, item) => { marker.lineWidth = currentLineWidth(item); });
  });

  el.viewport.addEventListener("pointerdown", pointerDown);
  el.viewport.addEventListener("pointermove", pointerMove);
  el.viewport.addEventListener("pointerup", pointerUp);
  el.viewport.addEventListener("pointercancel", pointerCancel);
  el.viewport.addEventListener("wheel", (event) => {
    if (state.loadedId) {
      event.preventDefault();
      zoomAt(Math.exp(-event.deltaY * 0.0015), event.clientX, event.clientY);
    }
  }, { passive: false });

  el.undo.addEventListener("click", undo);
  el.redo.addEventListener("click", redo);
  el.delete.addEventListener("click", deleteSelected);
  el.clear.addEventListener("click", clearMarkers);
  el.copy.addEventListener("click", copyAnnotatedImage);
  el.export.addEventListener("click", () => exportPng(false));
  el.exportNext.addEventListener("click", () => exportPng(true));
  el.zoomIn.addEventListener("click", () => zoomAt(1.2));
  el.zoomOut.addEventListener("click", () => zoomAt(1 / 1.2));
  el.fit.addEventListener("click", fitToView);
  $("#actualSizeButton").addEventListener("click", () => zoomAt(1 / state.scale));
  el.splitCountInput.addEventListener("change", () => setSplitCount(el.splitCountInput.value));
  window.addEventListener("keydown", handleKeyDown);
  window.addEventListener("keyup", (event) => {
    if (event.code === "Space") {
      state.spacePan = false;
      el.viewport.classList.remove("space-pan");
    }
  });
  window.addEventListener("resize", () => { if (state.loadedId) fitToView(); });

  runSelfCheck();
  setSplitCount(1);
  updateControls();
})();
