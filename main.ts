import { FileView, Notice, Plugin, TFile, WorkspaceLeaf } from "obsidian";
import * as docxPreview from "docx-preview";
import { init as initPptxPreview } from "pptx-preview";
import * as XLSX from "xlsx";

const VIEW_TYPE = "office-preview-view";

const EXTENSIONS = ["docx", "doc", "pptx", "ppt", "xlsx", "xls"];

const ZOOM_STEP = 0.1;
const MIN_ZOOM = 1.0;
const MAX_ZOOM = 3.0;

const MAX_FILE_MB = 50;

class OfficePreviewView extends FileView {
  private previewContainer: HTMLElement | null = null;
  private zoomSpacer: HTMLElement | null = null;
  private zoomTarget: HTMLElement | null = null;
  private zoomLevel = 1;
  private zoomLabelEl: HTMLElement | null = null;
  private naturalW = 0;
  private naturalH = 0;
  private pptxPreviewer: any = null;
  private renderId = 0;
  private toastTimer: number | null = null;
  private useTransformZoom = false; // true for pptx (GPU), false for docx/xlsx (zoom)

  getViewType(): string { return VIEW_TYPE; }
  getDisplayText(): string { return this.file?.name ?? "Office Preview"; }
  getIcon(): string { return "file-text"; }

  async onOpen(): Promise<void> {
    this.contentEl.empty();
    this.contentEl.style.cssText =
      "display:flex;flex-direction:column;height:100%;overflow:hidden";

    this.buildToolbar();
    this.previewContainer = this.contentEl.createDiv();
    this.previewContainer.className = "op-preview-container";

    this.zoomSpacer = this.previewContainer.createDiv();
    this.zoomSpacer.className = "op-zoom-spacer";

    this.zoomTarget = this.previewContainer.createDiv();
    this.zoomTarget.className = "op-zoom-target";

    this.attachZoomWheel();

    if (this.file) {
      await this.loadAndRender();
    } else {
      this.showPlaceholder();
    }
  }

  async onClose(): Promise<void> {
    this.clearToast();
    this.destroyPptxPreviewer();
    this.previewContainer = null;
    this.zoomSpacer = null;
    this.zoomTarget = null;
  }

  async onLoadFile(_file: TFile): Promise<void> { await this.loadAndRender(); }
  async onUnloadFile(_file: TFile): Promise<void> { this.destroyPptxPreviewer(); }

  private async loadAndRender(): Promise<void> {
    if (!this.file || !this.previewContainer || !this.zoomTarget || !this.zoomSpacer) return;

    if (this.file.stat.size > MAX_FILE_MB * 1024 * 1024) {
      new Notice(`文件较大 (${(this.file.stat.size / 1024 / 1024).toFixed(1)}MB)，渲染可能需要一些时间`);
    }

    const rid = ++this.renderId;
    this.destroyPptxPreviewer();

    // Reset zoom state for new file
    this.useTransformZoom = false;
    this.zoomTarget.style.position = "";
    this.zoomTarget.style.left = "";
    this.zoomTarget.style.top = "";
    this.zoomTarget.style.transform = "";
    this.zoomSpacer.style.width = "";
    this.zoomSpacer.style.height = "";
    this.naturalW = 0;
    this.naturalH = 0;
    this.zoomTarget.empty();
    (this.previewContainer.style as any).zoom = "";
    this.setZoom(1);
    this.showLoading();

    try {
      const buffer = await this.app.vault.readBinary(this.file);
      if (rid !== this.renderId) return;

      const ext = this.file.extension;
      if (ext === "pptx" || ext === "ppt") {
        this.useTransformZoom = true;
        this.zoomTarget.style.position = "absolute";
        this.zoomTarget.style.left = "0";
        this.zoomTarget.style.top = "0";
      }

      this.zoomTarget.empty();
      await this.dispatchRender(ext, buffer);

      if (this.useTransformZoom) {
        this.captureNaturalSize();
      }
    } catch (e: unknown) {
      if (rid !== this.renderId) return;
      this.showError(e instanceof Error ? e : new Error(String(e)));
    }
  }

  private captureNaturalSize(): void {
    if (!this.zoomTarget || !this.zoomSpacer) return;
    this.naturalW = this.zoomTarget.scrollWidth;
    this.naturalH = this.zoomTarget.scrollHeight;
    this.zoomSpacer.style.width = `${this.naturalW}px`;
    this.zoomSpacer.style.height = `${this.naturalH}px`;
  }

  private dispatchRender(ext: string, buffer: ArrayBuffer): Promise<void> | void {
    if (ext === "docx" || ext === "doc") return this.renderDocx(buffer);
    if (ext === "pptx" || ext === "ppt") return this.renderPptx(buffer);
    if (ext === "xlsx" || ext === "xls") return this.renderXlsx(buffer);
  }

  // ─── Toolbar ────────────────────────────────────────────

  private buildToolbar(): void {
    const bar = this.contentEl.createDiv();
    bar.className = "op-toolbar";

    const group = bar.createDiv();
    group.className = "op-toolbar-group";

    this.addBtn(group, "−", "缩小 (Ctrl+滚轮)", () => this.adjustZoom(-ZOOM_STEP));
    this.zoomLabelEl = group.createEl("span");
    this.zoomLabelEl.className = "op-toolbar-label";
    this.zoomLabelEl.textContent = "100%";
    this.addBtn(group, "+", "放大 (Ctrl+滚轮)", () => this.adjustZoom(ZOOM_STEP));
    this.addBtn(group, "1:1", "重置缩放", () => this.setZoom(1));
    this.addBtn(bar, "复制全文", "复制预览中的全部文本", () => this.copyAllText(), "op-toolbar-copy");
  }

  private addBtn(
    parent: HTMLElement, text: string, title: string, onClick: () => void, extraClass?: string
  ): void {
    const btn = parent.createEl("button");
    btn.className = "op-toolbar-btn" + (extraClass ? " " + extraClass : "");
    btn.textContent = text;
    btn.title = title;
    btn.addEventListener("click", onClick);
  }

  private attachZoomWheel(): void {
    this.contentEl.addEventListener("wheel", (e: WheelEvent) => {
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        this.adjustZoom(-Math.sign(e.deltaY) * ZOOM_STEP);
      }
    }, { passive: false });
  }

  private adjustZoom(delta: number): void { this.setZoom(this.zoomLevel + delta); }

  // ─── Zoom ────────────────────────────────────────────────

  private setZoom(level: number): void {
    const oldZoom = this.zoomLevel;
    this.zoomLevel = Math.round(Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, level)) * 100) / 100;

    if (this.useTransformZoom) {
      this.setTransformZoom(oldZoom);
    } else {
      this.setCssZoom();
    }

    if (this.zoomLabelEl) {
      this.zoomLabelEl.textContent = `${Math.round(this.zoomLevel * 100)}%`;
    }
  }

  private setCssZoom(): void {
    if (!this.previewContainer) return;
    (this.previewContainer.style as any).zoom =
      this.zoomLevel === 1 ? "" : String(this.zoomLevel);
  }

  private setTransformZoom(oldZoom: number): void {
    if (!this.zoomTarget || !this.zoomSpacer || !this.previewContainer) return;

    const pc = this.previewContainer;
    const vw = pc.clientWidth;
    const vh = pc.clientHeight;
    const nw = this.naturalW || vw;
    const nh = this.naturalH || vh;

    const cx = (pc.scrollLeft + vw / 2) / (nw * oldZoom || 1);
    const cy = (pc.scrollTop + vh / 2) / (nh * oldZoom || 1);

    this.zoomSpacer.style.width = `${nw * this.zoomLevel}px`;
    this.zoomSpacer.style.height = `${nh * this.zoomLevel}px`;

    if (this.zoomLevel === 1) {
      this.zoomTarget.style.transform = "";
    } else {
      this.zoomTarget.style.transform = `scale(${this.zoomLevel})`;
      this.zoomTarget.style.transformOrigin = "0 0";
    }

    pc.scrollLeft = cx * nw * this.zoomLevel - vw / 2;
    pc.scrollTop = cy * nh * this.zoomLevel - vh / 2;
  }

  // ─── DOCX ───────────────────────────────────────────────

  private async renderDocx(buffer: ArrayBuffer): Promise<void> {
    if (!this.zoomTarget) return;
    await docxPreview.renderAsync(buffer, this.zoomTarget, undefined, {
      inWrapper: true,
      useBase64URL: true,
    });
  }

  // ─── PPTX ───────────────────────────────────────────────

  private async renderPptx(buffer: ArrayBuffer): Promise<void> {
    if (!this.zoomTarget) return;
    const pw = this.previewContainer?.clientWidth || 960;
    this.pptxPreviewer = initPptxPreview(this.zoomTarget, {
      width: pw, mode: "list",
    });
    await this.pptxPreviewer.preview(buffer);
  }

  private destroyPptxPreviewer(): void {
    if (!this.pptxPreviewer) return;
    try { this.pptxPreviewer.destroy(); } catch { /* ignore */ }
    this.pptxPreviewer = null;
  }

  // ─── XLSX ───────────────────────────────────────────────

  private renderXlsx(buffer: ArrayBuffer): void {
    if (!this.zoomTarget) return;
    const wb = XLSX.read(buffer, { type: "array" });
    if (wb.SheetNames.length === 0) {
      this.zoomTarget.createDiv({ text: "工作簿中无工作表" });
      return;
    }

    const wrapper = this.zoomTarget.createDiv();
    wrapper.className = "op-xlsx-wrapper";

    let current = wb.SheetNames[0];
    let area: HTMLElement;

    if (wb.SheetNames.length > 1) {
      const tabs = wrapper.createDiv();
      tabs.className = "op-xlsx-tabs";
      wb.SheetNames.forEach((n, i) => {
        const tab = tabs.createEl("button");
        tab.className = "op-xlsx-tab";
        tab.textContent = n;
        if (i === 0) tab.classList.add("op-xlsx-tab-active");
        tab.addEventListener("click", () => {
          tabs.querySelectorAll(".op-xlsx-tab").forEach((t) =>
            t.classList.remove("op-xlsx-tab-active"));
          tab.classList.add("op-xlsx-tab-active");
          area.empty();
          this.renderSheet(wb, n, area);
        });
      });
    }

    area = wrapper.createDiv();
    area.className = "op-xlsx-area";
    this.renderSheet(wb, current, area);
  }

  private renderSheet(wb: XLSX.WorkBook, name: string, container: HTMLElement): void {
    const sheet = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1, defval: "", raw: false,
    });

    if (data.length === 0) {
      container.createDiv({ text: "工作表为空", cls: "op-xlsx-empty" });
      return;
    }

    const maxRows = Math.min(200, data.length);
    const maxCols = data.reduce((max, r) => Math.max(max, r.length), 0);
    const display = data.slice(0, maxRows);

    const tw = container.createDiv();
    tw.className = "op-xlsx-table-wrapper";
    const table = tw.createEl("table");
    table.className = "op-xlsx-table";

    display.forEach((row, ri) => {
      const tr = table.createEl("tr");
      if (ri === 0) tr.classList.add("op-xlsx-header-row");
      const rn = tr.createEl("td");
      rn.className = "op-xlsx-row-num";
      rn.textContent = String(ri + 1);
      for (let c = 0; c < maxCols; c++) {
        const td = tr.createEl("td");
        const v = row[c];
        td.textContent = v != null ? String(v) : "";
      }
    });

    const info = container.createDiv();
    info.className = "op-xlsx-info";
    info.textContent = `显示 ${display.length} × ${maxCols}（共 ${data.length} 行）`;
  }

  // ─── Helpers ────────────────────────────────────────────

  private showLoading(): void {
    if (!this.zoomTarget) return;
    this.zoomTarget.createDiv({
      cls: "op-loading",
    }).innerHTML = `<div class="op-spinner"></div><div>加载预览中...</div>`;
  }

  private showError(e: Error): void {
    if (!this.zoomTarget) return;
    this.zoomTarget.empty();
    this.zoomTarget.createDiv({
      cls: "op-error",
    }).innerHTML = `<div class="op-error-icon">!</div><div>预览失败: ${e.message || "未知错误"}</div>`;
  }

  private showPlaceholder(): void {
    if (!this.zoomTarget) return;
    this.zoomTarget.createDiv({
      text: "点击文件列表中的 Office 文件进行预览", cls: "op-placeholder",
    });
  }

  private copyAllText(): void {
    if (!this.zoomTarget) return;
    const text = this.zoomTarget.innerText?.trim();
    if (!text) return;

    const done = () => this.showCopyToast("已复制到剪贴板");
    const fallback = () => {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.cssText = "position:fixed;left:-9999px";
      ta.select();
      document.execCommand("copy");
      ta.remove();
      done();
    };
    navigator.clipboard?.writeText(text).then(done, fallback) ?? fallback();
  }

  private showCopyToast(msg: string): void {
    this.clearToast();
    const toast = this.contentEl.createDiv();
    toast.className = "op-toast";
    toast.textContent = msg;
    this.toastTimer = window.setTimeout(() => this.clearToast(), 2000);
  }

  private clearToast(): void {
    if (this.toastTimer != null) { clearTimeout(this.toastTimer); this.toastTimer = null; }
    this.contentEl.querySelectorAll(".op-toast").forEach((t) => t.remove());
  }
}

export default class OfficePreviewPlugin extends Plugin {
  async onload(): Promise<void> {
    this.registerView(VIEW_TYPE, (leaf) => new OfficePreviewView(leaf));
    this.registerExtensions(EXTENSIONS, VIEW_TYPE);

    this.registerEvent(
      this.app.workspace.on("file-menu", (menu, file) => {
        if (file instanceof TFile && EXTENSIONS.includes(file.extension)) {
          menu.addItem((item) =>
            item.setTitle("预览此文件").setIcon("file-text").onClick(() => this.openPreview(file))
          );
        }
      })
    );

    this.addCommand({
      id: "preview-current-file",
      name: "预览当前文件",
      checkCallback: (checking) => {
        const f = this.app.workspace.getActiveFile();
        if (!f || !EXTENSIONS.includes(f.extension)) return false;
        if (!checking) this.openPreview(f);
        return true;
      },
    });
  }

  async onunload(): Promise<void> {
    this.app.workspace.detachLeavesOfType(VIEW_TYPE);
  }

  private async openPreview(file: TFile): Promise<void> {
    const leaves = this.app.workspace.getLeavesOfType(VIEW_TYPE);
    const leaf = leaves.length > 0 ? leaves[0] : this.app.workspace.getLeaf(true);
    await leaf.openFile(file);
    this.app.workspace.revealLeaf(leaf);
  }
}
