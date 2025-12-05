/* eslint-disable @typescript-eslint/no-explicit-any,
                  @typescript-eslint/no-inferrable-types,
                  @typescript-eslint/no-empty-function */

import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { renderAsync } from "docx-preview";
import * as pdfjsLib from "pdfjs-dist";

export class DocumentPreviewInline implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;
    private container: HTMLDivElement;

    // Inline preview elements
    private previewContainer: HTMLDivElement | null = null;
    private previewToolbarTitle: HTMLSpanElement | null = null;
    private previewToolbarActions: HTMLDivElement | null = null;
    private zoomWrapper: HTMLDivElement | null = null;
    private zoomLevel: number = 1;

    private currentFileUrl: string | null = null;
    private currentFileName: string | null = null;
    private currentMimeType: string | null = null;
    private currentDownloadUrl: string | null = null;

    // Worker local de pdf.js
    private pdfWorkerUrl: string | null = null;
    private pdfWorkerUrlPromise: Promise<string> | null = null;

    constructor() { }

    // -------------------------------------------------------------------------
    // Control lifecycle
    // -------------------------------------------------------------------------

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;

        // Configurar pdf.js para usar el worker local (no CDN)
        this.ensurePdfWorkerUrl()
            .then((workerUrl) => {
                (pdfjsLib as any).GlobalWorkerOptions.workerSrc = workerUrl;
            })
            .catch((err) => {
                console.error("Error setting pdf.js workerSrc", err);
            });

        // Root = full preview panel (no button)
        const root = document.createElement("div");
        Object.assign(root.style, {
            width: "100%",
            boxSizing: "border-box"
        });

        // Preview container
        this.previewContainer = document.createElement("div");
        this.previewContainer.className = "gwm-doc-preview";
        Object.assign(this.previewContainer.style, {
            border: "1px solid #e1e1e1",
            borderRadius: "4px",
            backgroundColor: "#f5f5f5",
            minHeight: this.getViewerHeight(),
            maxHeight: this.getViewerHeight(),
            overflow: "auto",
            padding: "8px 12px",
            boxSizing: "border-box"
        });

        // Toolbar (title + actions)
        const toolbar = document.createElement("div");
        Object.assign(toolbar.style, {
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            marginBottom: "8px",
            fontSize: "12px"
        });

        this.previewToolbarTitle = document.createElement("span");
        this.previewToolbarTitle.innerText = "Document preview";

        this.previewToolbarActions = document.createElement("div");
        Object.assign(this.previewToolbarActions.style, {
            display: "flex",
            gap: "6px"
        });

        // Zoom controls
        const zOut = this.makeButton("−", () => this.applyZoom(this.zoomLevel - 0.1));
        const zIn = this.makeButton("+", () => this.applyZoom(this.zoomLevel + 0.1));
        const zReset = this.makeButton("100%", () => this.applyZoom(1));

        this.previewToolbarActions.appendChild(zOut);
        this.previewToolbarActions.appendChild(zIn);
        this.previewToolbarActions.appendChild(zReset);

        // Optional download button
        if (this.context.parameters.EnableDownload.raw) {
            const downloadBtn = this.makeButton("Download", () => this.downloadCurrentFile(), true);
            this.previewToolbarActions.appendChild(downloadBtn);
        }

        toolbar.appendChild(this.previewToolbarTitle);
        toolbar.appendChild(this.previewToolbarActions);

        // Zoom wrapper (content area)
        this.zoomWrapper = document.createElement("div");
        this.zoomWrapper.setAttribute("data-zoom-wrapper", "true");
        Object.assign(this.zoomWrapper.style, {
            transformOrigin: "top left",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: "16px",
            padding: "16px",
            boxSizing: "border-box",
            width: "100%"
        });

        this.previewContainer.appendChild(toolbar);
        this.previewContainer.appendChild(this.zoomWrapper);

        root.appendChild(this.previewContainer);
        this.container.appendChild(root);

        // Initial zoom
        this.applyZoom(1);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        // Every time the form refreshes, try to load the preview automatically
        void this.autoLoadPreview();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.cleanupFileUrl();

        if (this.pdfWorkerUrl) {
            URL.revokeObjectURL(this.pdfWorkerUrl);
            this.pdfWorkerUrl = null;
        }
    }

    private getViewerHeight(): string {
        const h = this.context?.parameters?.ViewerHeight?.raw;
        return h && h > 0 ? `${h}px` : "500px";
    }
// -------------------------------------------------------------------------
// Worker local de pdf.js
// -------------------------------------------------------------------------
private ensurePdfWorkerUrl(): Promise<string> {
    if (this.pdfWorkerUrlPromise) {
        return this.pdfWorkerUrlPromise;
    }

    this.pdfWorkerUrlPromise = new Promise<string>((resolve, reject) => {
        const workerPath = "pdf.worker.min.html"; 

        this.context.resources.getResource(
            workerPath,
            (data: string) => {
                try {
                    
                    const byteChars = atob(data);
                    const byteNumbers = new Array(byteChars.length);
                    for (let i = 0; i < byteChars.length; i++) {
                        byteNumbers[i] = byteChars.charCodeAt(i);
                    }
                    const byteArray = new Uint8Array(byteNumbers);

                    const blob = new Blob([byteArray], { type: "text/javascript" });
                    const url = URL.createObjectURL(blob);
                    this.pdfWorkerUrl = url;
                    resolve(url);
                } catch (e) {
                    reject(e);
                }
            },
            () => {
                
                reject(new Error("Error loading pdf.js worker resource"));
            }
        );
    });

    return this.pdfWorkerUrlPromise;
}



    // -------------------------------------------------------------------------
    // Supported types
    // -------------------------------------------------------------------------

    private readonly supportedExtensions: string[] = [
        "pdf",
        "doc",
        "docx",
        "png",
        "jpg",
        "jpeg",
        "gif"
    ];

    // -------------------------------------------------------------------------
    // Auto-load preview
    // -------------------------------------------------------------------------

    private async autoLoadPreview(): Promise<void> {
        if (!this.zoomWrapper) return;

        const entityId = (this.context.mode as any).contextInfo?.entityId;
        const entityLogicalName = this.context.parameters.EntityLogicalName.raw;
        const fileColumnName = this.context.parameters.FileColumnName.raw;

        // New record not saved yet
        if (!entityId) {
            this.setInfoMessage(
                "Save the record and upload a file to see the document preview."
            );
            return;
        }

        if (!entityLogicalName || !fileColumnName) {
            this.setInfoMessage(
                "Control is missing configuration (EntityLogicalName or FileColumnName)."
            );
            return;
        }

        try {
            const hasPreview = await this.hasPreviewableAttachment(
                entityLogicalName,
                entityId,
                fileColumnName
            );

            if (!hasPreview) {
                this.setInfoMessage("No previewable file in Attachment.");
                return;
            }

            // There is a supported file => download and render
            const file = await this.downloadFile(entityLogicalName, entityId, fileColumnName);

            this.currentFileName = file.fileName;
            this.currentMimeType = file.mimeType;

            this.cleanupFileUrl();
            this.currentFileUrl = URL.createObjectURL(file.blob);

            if (this.previewToolbarTitle) {
                this.previewToolbarTitle.innerText = this.currentFileName || "Document preview";
            }

            await this.renderFilePreview(file);
        } catch (err: any) {
            console.error("Error loading document preview", err);

            if (err.status === 403) {
                this.setInfoMessage("You do not have permission to view this document.");
            } else {
                this.setInfoMessage("Error loading document preview.");
            }
        }
    }

    private async hasPreviewableAttachment(
        entityLogicalName: string,
        entityId: string,
        fileColumnName: string
    ): Promise<boolean> {
        const webApi = this.context.webAPI;

        const nameColumn = `${fileColumnName}_name`;
        const select = `${fileColumnName},${nameColumn}`;

        const record = await webApi.retrieveRecord(entityLogicalName, entityId, `?$select=${select}`);

        const fileRef = (record as any)[fileColumnName];

        if (!fileRef) return false;

        const fileName = ((record as any)[nameColumn] as string) || "";
        if (!fileName) return true;

        const dot = fileName.lastIndexOf(".");
        if (dot <= 0 || dot === fileName.length - 1) return false;

        const ext = fileName.substring(dot + 1).toLowerCase();
        return this.supportedExtensions.includes(ext);
    }

    private setInfoMessage(text: string): void {
        if (!this.zoomWrapper) return;

        this.zoomWrapper.innerHTML = "";
        const block = document.createElement("div");
        block.innerText = text;
        Object.assign(block.style, {
            padding: "16px",
            backgroundColor: "#ffffff",
            borderRadius: "4px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
            fontSize: "12px",
            color: "#323130",
            maxWidth: "700px",
            textAlign: "center",
            width: "100%",
            boxSizing: "border-box"
        });

        this.zoomWrapper.appendChild(block);

        if (this.previewToolbarTitle) {
            this.previewToolbarTitle.innerText = "Document preview";
        }
    }

    // -------------------------------------------------------------------------
    // Buttons / zoom / download helpers
    // -------------------------------------------------------------------------

    private makeButton(label: string, handler: () => void, isPrimary = false): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.innerText = label;
        btn.onclick = handler;

        Object.assign(btn.style, {
            cursor: "pointer",
            padding: "4px 10px",
            fontSize: "12px",
            borderRadius: "4px",
            border: isPrimary ? "1px solid #106ebe" : "1px solid #8a8886",
            backgroundColor: isPrimary ? "#0078d4" : "#ffffff",
            color: isPrimary ? "#ffffff" : "#323130",
            minWidth: "48px",
            transition: "background-color 0.15s ease, border-color 0.15s ease"
        });

        btn.onmouseover = () => {
            btn.style.backgroundColor = isPrimary ? "#106ebe" : "#f3f2f1";
        };

        btn.onmouseout = () => {
            btn.style.backgroundColor = isPrimary ? "#0078d4" : "#ffffff";
        };

        return btn;
    }

    private applyZoom(z: number): void {
        if (!this.zoomWrapper) return;
        this.zoomLevel = Math.min(Math.max(z, 0.5), 3);
        this.zoomWrapper.style.transform = `scale(${this.zoomLevel})`;
    }

    private downloadCurrentFile(): void {
        if (this.currentDownloadUrl) {
            const link = document.createElement("a");
            link.href = this.currentDownloadUrl;
            link.download = this.currentFileName || "file";
            link.target = "_blank";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            return;
        }

        if (this.currentFileUrl) {
            const link = document.createElement("a");
            link.href = this.currentFileUrl;
            link.download = this.currentFileName || "file";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            return;
        }
    }

    private cleanupFileUrl(): void {
        if (this.currentFileUrl) {
            URL.revokeObjectURL(this.currentFileUrl);
            this.currentFileUrl = null;
        }
    }

    // -------------------------------------------------------------------------
    // Dataverse Web API helpers
    // -------------------------------------------------------------------------

    private async getFileUrl(
        _entityLogicalName: string,
        entityId: string,
        fileColumnName: string
    ): Promise<string> {
        const entityPluralName = this.context.parameters.EntityPluralName.raw;
        const cleanId = (entityId || "").replace(/[{}]/g, "");

        const orgUrl =
            (this.context as any).page?.getClientUrl?.() ||
            (window as any).Xrm?.Utility?.getGlobalContext?.().getClientUrl?.() ||
            window.location.origin;

        const normalizedColumn = (fileColumnName || "").trim();

        const url = `${orgUrl}/api/data/v9.0/${entityPluralName}(${cleanId})/${normalizedColumn}`;
        return url;
    }

    private async downloadFile(
        entityLogicalName: string,
        entityId: string,
        fileColumnName: string
    ): Promise<{ blob: Blob; fileName: string; mimeType: string }> {
        const webApi = this.context.webAPI;

        const nameColumn = `${fileColumnName}_name`;
        const select = `${fileColumnName},${nameColumn}`;
        const record = await webApi.retrieveRecord(entityLogicalName, entityId, `?$select=${select}`);

        const fileName = ((record as any)[nameColumn] as string) || "document";

        const fileUrl = await this.getFileUrl(entityLogicalName, entityId, fileColumnName);

        // Dataverse $value endpoint for direct download
        this.currentDownloadUrl = `${fileUrl}/$value`;

        const response = await fetch(fileUrl, {
            method: "GET",
            headers: {
                Accept: "application/json"
            }
        });

        if (!response.ok) {
            const error: any = new Error("Error retrieving file");
            error.status = response.status;
            throw error;
        }

        const json = (await response.json()) as { value: string };
        const base64 = json.value;

        const byteString = atob(base64);
        const len = byteString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = byteString.charCodeAt(i);
        }

        const lowerName = fileName.toLowerCase();
        let mimeType = "application/octet-stream";

        if (lowerName.endsWith(".pdf")) {
            mimeType = "application/pdf";
        } else if (lowerName.endsWith(".docx")) {
            mimeType =
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        } else if (lowerName.endsWith(".doc")) {
            mimeType = "application/msword";
        } else if (lowerName.endsWith(".png")) {
            mimeType = "image/png";
        } else if (lowerName.endsWith(".jpg") || lowerName.endsWith(".jpeg")) {
            mimeType = "image/jpeg";
        } else if (lowerName.endsWith(".gif")) {
            mimeType = "image/gif";
        }

        const blob = new Blob([bytes], { type: mimeType });

        return {
            blob,
            fileName,
            mimeType
        };
    }

    // -------------------------------------------------------------------------
    // Rendering (PDF / image / docx / fallback)
    // -------------------------------------------------------------------------

    private async renderFilePreview(file: { blob: Blob; fileName: string; mimeType: string }): Promise<void> {
        if (!this.zoomWrapper) return;

        const zoomWrapper = this.zoomWrapper;
        zoomWrapper.innerHTML = "";

        Object.assign(zoomWrapper.style, {
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: "16px",
            padding: "16px",
            boxSizing: "border-box",
            width: "100%"
        });

        const lower = file.fileName.toLowerCase();
        const mime = file.mimeType.toLowerCase();

        const createInfoBlock = (text: string): HTMLDivElement => {
            const div = document.createElement("div");
            div.innerText = text;
            Object.assign(div.style, {
                padding: "16px",
                backgroundColor: "#ffffff",
                borderRadius: "4px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                fontSize: "12px",
                color: "#323130",
                maxWidth: "700px",
                textAlign: "center",
                width: "100%",
                boxSizing: "border-box"
            });
            return div;
        };

        // ---------------------------------------------------------------------
        // PDF 
        // ---------------------------------------------------------------------
        if (mime === "application/pdf" || lower.endsWith(".pdf")) {
            const loading = createInfoBlock("Loading PDF preview...");
            zoomWrapper.appendChild(loading);

            try {

                await this.ensurePdfWorkerUrl().catch(() => { });

                const arrayBuffer = await file.blob.arrayBuffer();
                const pdf = await (pdfjsLib as any).getDocument({ data: arrayBuffer }).promise;

                zoomWrapper.innerHTML = "";

                // Available width inside the preview (minus padding)
                const containerWidth =
                    zoomWrapper.clientWidth ||
                    this.previewContainer?.clientWidth ||
                    800;
                const renderWidth = containerWidth - 32;

                const numPages = pdf.numPages;
                for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                    const page = await pdf.getPage(pageNum);

                    // Base viewport to know original width
                    const baseViewport = page.getViewport({ scale: 1 });
                    const scale = renderWidth / baseViewport.width; // fit to width
                    const viewport = page.getViewport({ scale });

                    const canvas = document.createElement("canvas");
                    const context2d = canvas.getContext("2d")!;
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;

                    Object.assign(canvas.style, {
                        display: "block",
                        background: "#ffffff",
                        borderRadius: "4px",
                        width: "100%",
                        height: "auto"
                    });

                    const pageContainer = document.createElement("div");
                    Object.assign(pageContainer.style, {
                        backgroundColor: "#ffffff",
                        padding: "12px",
                        borderRadius: "6px",
                        boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                        margin: "0 auto",
                        width: "100%",
                        maxWidth: "100%",
                        boxSizing: "border-box"
                    });

                    pageContainer.appendChild(canvas);
                    zoomWrapper.appendChild(pageContainer);

                    const renderContext = {
                        canvasContext: context2d,
                        viewport: viewport
                    };

                    await page.render(renderContext).promise;
                }
            } catch (err) {
                console.error("Error rendering PDF with pdf.js", err);
                zoomWrapper.innerHTML = "";
                zoomWrapper.appendChild(
                    createInfoBlock(
                        "Could not render PDF preview. You can still use the Download button to open the file."
                    )
                );
            }

            return;
        }

        // ---------------------------------------------------------------------
        // IMAGE 
        // ---------------------------------------------------------------------
        if (mime.startsWith("image/")) {
            const loading = createInfoBlock("Loading image preview...");
            zoomWrapper.appendChild(loading);

            const reader = new FileReader();

            reader.onload = () => {
                zoomWrapper.innerHTML = "";

                const container = document.createElement("div");
                Object.assign(container.style, {
                    backgroundColor: "#ffffff",
                    padding: "12px",
                    borderRadius: "6px",
                    boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                    maxWidth: "900px",
                    width: "100%",
                    display: "flex",
                    justifyContent: "center",
                    boxSizing: "border-box"
                });

                const img = document.createElement("img");
                img.src = reader.result as string;
                Object.assign(img.style, {
                    maxWidth: "100%",
                    height: "auto",
                    display: "block",
                    borderRadius: "4px"
                });

                container.appendChild(img);
                zoomWrapper.appendChild(container);
            };

            reader.onerror = () => {
                zoomWrapper.innerHTML = "";
                zoomWrapper.appendChild(
                    createInfoBlock(
                        "Could not render image preview. You can still use the Download button to open the file."
                    )
                );
            };

            reader.readAsDataURL(file.blob);

            return;
        }

        // ---------------------------------------------------------------------
        // DOCX  
        // ---------------------------------------------------------------------
        if (lower.endsWith(".docx")) {
            const wordContainerOuter = document.createElement("div");
            Object.assign(wordContainerOuter.style, {
                backgroundColor: "#f5f5f5",
                padding: "12px",
                borderRadius: "6px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                maxWidth: "900px",
                width: "100%",
                boxSizing: "border-box",
                overflowX: "hidden"
            });

            const wordInner = document.createElement("div");
            Object.assign(wordInner.style, {
                backgroundColor: "#ffffff",
                padding: "16px",
                borderRadius: "4px",
                minHeight: "200px",
                width: "100%",
                maxWidth: "100%",
                boxSizing: "border-box",
                overflowX: "hidden"
            });

            wordContainerOuter.appendChild(wordInner);
            zoomWrapper.appendChild(wordContainerOuter);

            // Render DOCX
            await renderAsync(file.blob, wordInner, undefined, {
                className: "docx-wrapper",
                inWrapper: true
            });

            // Adjust styles for better fit
            const wrappers = wordInner.querySelectorAll<HTMLElement>(".docx-wrapper, section.docx-wrapper");
            wrappers.forEach(w => {
                w.style.boxSizing = "border-box";
                w.style.maxWidth = "100%";
                w.style.width = "100%";
                w.style.margin = "0 auto";
                w.style.padding = "50px";
            });

            const docxEls = wordInner.querySelectorAll<HTMLElement>(".docx");
            docxEls.forEach(el => {
                el.style.maxWidth = "100%";
                el.style.width = "100%";
                el.style.boxSizing = "border-box";
                el.style.margin = "0 auto";
            });

            return;
        }

        // Legacy .DOC
        if (lower.endsWith(".doc")) {
            const msg = createInfoBlock(
                "This file is a legacy .doc format which cannot be rendered inline. Click the button below to open it in a new tab."
            );

            const btn = this.makeButton(
                "Open document",
                () => {
                    if (this.currentFileUrl) {
                        window.open(this.currentFileUrl, "_blank");
                    }
                },
                true
            );

            const actions = document.createElement("div");
            Object.assign(actions.style, {
                marginTop: "8px",
                display: "flex",
                justifyContent: "center"
            });

            actions.appendChild(btn);

            const container = document.createElement("div");
            Object.assign(container.style, {
                display: "flex",
                flexDirection: "column",
                alignItems: "center"
            });

            container.appendChild(msg);
            container.appendChild(actions);

            zoomWrapper.appendChild(container);
            return;
        }

        // Fallback
        zoomWrapper.appendChild(
            createInfoBlock(`Preview is not available for this file type (${file.mimeType}).`)
        );
    }
}
