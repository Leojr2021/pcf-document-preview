/* eslint-disable @typescript-eslint/no-explicit-any,
                  @typescript-eslint/no-inferrable-types,
                  @typescript-eslint/no-empty-function */

import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { renderAsync } from "docx-preview";
import * as pdfjsLib from "pdfjs-dist";

const pdfjsVersion = (pdfjsLib as any).version || "5.4.449";

// Configure pdf.js worker from CDN (required for rendering PDFs)
(pdfjsLib as any).GlobalWorkerOptions.workerSrc =
    `https://cdn.jsdelivr.net/npm/pdfjs-dist@${pdfjsVersion}/build/pdf.worker.min.mjs`;

export class DocumentPreview implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private context: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged: () => void;
    private container: HTMLDivElement;

    private previewButton: HTMLButtonElement;
    private fileLabel: HTMLSpanElement;

    private modalOverlay: HTMLDivElement | null = null;
    private modalContent: HTMLDivElement | null = null;
    private modalBody: HTMLDivElement | null = null;
    private zoomLevel: number = 1;

    private currentFileUrl: string | null = null;
    private currentFileName: string | null = null;
    private currentMimeType: string | null = null;
    private currentDownloadUrl: string | null = null;

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

        // Wrapper for button + status label
        const wrapper = document.createElement("div");
        wrapper.style.display = "flex";
        wrapper.style.alignItems = "center";
        wrapper.style.gap = "8px";

        // Main "Preview document" button (simple style, model-driven friendly)
        this.previewButton = document.createElement("button");
        this.previewButton.innerText = "Preview document";
        this.previewButton.onclick = () => this.onPreviewClick();
        this.previewButton.style.padding = "4px 8px";
        this.previewButton.style.fontSize = "12px";
        this.previewButton.style.cursor = "pointer";

        // Status label (shows messages like "Checking attachment...")
        this.fileLabel = document.createElement("span");
        this.fileLabel.style.fontSize = "12px";
        this.fileLabel.style.color = "#666";

        wrapper.appendChild(this.previewButton);
        wrapper.appendChild(this.fileLabel);
        this.container.appendChild(wrapper);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;

        this.fileLabel.innerText = "Checking attachment...";
        this.previewButton.disabled = true;

        void this.checkAttachmentAndUpdateUI();
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this.cleanupFileUrl();
        this.destroyModal();
    }

    // -------------------------------------------------------------------------
    // Supported file types (extensions-based)
    // -------------------------------------------------------------------------

    private readonly supportedMimeTypes: string[] = [
        "application/pdf",
        "application/msword",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "image/png",
        "image/jpeg",
        "image/gif"
    ];

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
    // Attachment checking / enabling preview button
    // -------------------------------------------------------------------------

    /**
     * Checks if the configured file column on the current record
     * has a previewable file (based on extension).
     */
    private async checkAttachmentAndUpdateUI(): Promise<void> {
        const entityId = (this.context.mode as any).contextInfo?.entityId;
        const entityLogicalName = this.context.parameters.EntityLogicalName.raw;
        const fileColumnName = this.context.parameters.FileColumnName.raw;

        if (!entityId || !entityLogicalName || !fileColumnName) {
            this.fileLabel.innerText =
                "Control is missing configuration (EntityLogicalName or FileColumnName).";
            this.previewButton.disabled = true;
            return;
        }

        try {
            const hasPreview = await this.hasPreviewableAttachment(
                entityLogicalName,
                entityId,
                fileColumnName
            );

            if (hasPreview) {
                this.fileLabel.innerText = "";
                this.previewButton.disabled = false;
            } else {
                this.fileLabel.innerText = "No previewable file in Attachment.";
                this.previewButton.disabled = true;
            }
        } catch (e) {
            console.error("Error checking attachment", e);
            this.fileLabel.innerText = "Could not check Attachment.";
            this.previewButton.disabled = true;
        }
    }

    /**
     * Uses Web API to retrieve the file column and its name, then decides
     * if the extension is supported for inline preview.
     */
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

        if (!fileRef) {
            // No file value at all
            return false;
        }

        const fileName = ((record as any)[nameColumn] as string) || "";

        if (!fileName) {
            // There is a file but no name; allow preview as a fallback
            return true;
        }

        const dot = fileName.lastIndexOf(".");
        if (dot <= 0 || dot === fileName.length - 1) {
            // No usable extension
            return false;
        }

        const ext = fileName.substring(dot + 1).toLowerCase();
        return this.supportedExtensions.includes(ext);
    }

    // -------------------------------------------------------------------------
    // Preview button handler
    // -------------------------------------------------------------------------

    private async onPreviewClick(): Promise<void> {
        if (this.previewButton.disabled) {
            return;
        }

        const entityId = (this.context.mode as any).contextInfo?.entityId;
        const entityLogicalName = this.context.parameters.EntityLogicalName.raw;
        const fileColumnName = this.context.parameters.FileColumnName.raw;

        if (!entityId || !entityLogicalName || !fileColumnName) {
            alert("Control is missing configuration (EntityLogicalName or FileColumnName).");
            return;
        }

        try {
            // Download file bytes + metadata
            const file = await this.downloadFile(entityLogicalName, entityId, fileColumnName);

            this.currentFileName = file.fileName;
            this.currentMimeType = file.mimeType;

            // Manage object URL lifecycle
            this.cleanupFileUrl();
            this.currentFileUrl = URL.createObjectURL(file.blob);

            // Open modal and render the file inside
            this.openModal();
            await this.renderFilePreview(file);
        } catch (err: any) {
            console.error(err);

            if (err.status === 403) {
                alert("You do not have permission to view this document.");
            } else {
                alert("Error loading document preview.");
            }
        }
    }

    // -------------------------------------------------------------------------
    // Modal UI (model-driven / Fluent-like look and feel)
    // -------------------------------------------------------------------------

    /**
     * Creates and shows the modal overlay + container with a Fluent-style header.
     */
    private openModal(): void {
        // Ensure any previous modal is removed first
        this.destroyModal();
        this.zoomLevel = 1;

        // Fullscreen overlay
        this.modalOverlay = document.createElement("div");
        Object.assign(this.modalOverlay.style, {
            position: "fixed",
            top: "0",
            left: "0",
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0,0,0,0.3)",
            zIndex: "9999",
            display: "flex",
            alignItems: "center",
            justifyContent: "center"
        });

        // Central modal container
        this.modalContent = document.createElement("div");
        Object.assign(this.modalContent.style, {
            backgroundColor: "#ffffff",
            borderRadius: "6px",
            boxShadow: "0 4px 12px rgba(0,0,0,0.2)",
            width: "80%",
            maxWidth: "1000px",
            maxHeight: this.context.parameters.ViewerHeight.raw
                ? `${this.context.parameters.ViewerHeight.raw}px`
                : "600px",
            display: "flex",
            flexDirection: "column",
            overflow: "hidden"
        });

        // Header (title + actions)
        const header = document.createElement("div");
        Object.assign(header.style, {
            padding: "10px 16px",
            borderBottom: "1px solid #e1e1e1",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            backgroundColor: "#f3f2f1"
        });

        const title = document.createElement("div");
        title.innerText = this.currentFileName || "Document preview";
        Object.assign(title.style, {
            fontWeight: "600",
            fontSize: "14px",
            color: "#323130"
        });

        const actions = document.createElement("div");
        Object.assign(actions.style, {
            display: "flex",
            gap: "6px"
        });

        // Zoom controls
        const zOut = this.makeButton("−", () => this.applyZoom(this.zoomLevel - 0.1));
        const zIn = this.makeButton("+", () => this.applyZoom(this.zoomLevel + 0.1));
        const zReset = this.makeButton("100%", () => this.applyZoom(1));

        actions.appendChild(zOut);
        actions.appendChild(zIn);
        actions.appendChild(zReset);

        // Optional "Download" button (primary style)
        if (this.context.parameters.EnableDownload.raw) {
            actions.appendChild(
                this.makeButton("Download", () => {
                    // Prefer direct download URL (Dataverse $value endpoint)
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

                    // Fallback: use object URL if available
                    if (this.currentFileUrl) {
                        const link = document.createElement("a");
                        link.href = this.currentFileUrl;
                        link.download = this.currentFileName || "file";
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                        return;
                    }

                    console.warn("No file URL available for download.");
                }, true)
            );
        }

        // Close button
        actions.appendChild(this.makeButton("Close", () => this.destroyModal()));

        header.appendChild(title);
        header.appendChild(actions);

        // Modal body area (scrollable preview region)
        this.modalBody = document.createElement("div");
        Object.assign(this.modalBody.style, {
            flex: "1",
            overflow: "auto",
            backgroundColor: "#f5f5f5",
            padding: "12px"
        });

        // Inner wrapper where PDF pages / images / docx content will be rendered
        const zoomWrapper = document.createElement("div");
        zoomWrapper.setAttribute("data-zoom-wrapper", "true");
        Object.assign(zoomWrapper.style, {
            transformOrigin: "top left",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: "16px",
            padding: "16px",
            boxSizing: "border-box"
        });

        this.modalBody.appendChild(zoomWrapper);

        this.modalContent.appendChild(header);
        this.modalContent.appendChild(this.modalBody);

        this.modalOverlay.appendChild(this.modalContent);
        document.body.appendChild(this.modalOverlay);

        // Initialize zoom to 100%
        this.applyZoom(1);
    }

    /**
     * Creates a Fluent-like button (primary or secondary) for the modal header.
     */
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
            minWidth: "64px",
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

    /**
     * Removes the modal from DOM and cleans related references.
     */
    private destroyModal(): void {
        if (this.modalOverlay && this.modalOverlay.parentNode) {
            this.modalOverlay.parentNode.removeChild(this.modalOverlay);
        }
        this.modalOverlay = null;
        this.modalContent = null;
        this.modalBody = null;
        this.cleanupFileUrl();
    }

    /**
     * Applies a CSS transform-based zoom to the inner preview wrapper.
     */
    private applyZoom(z: number): void {
        if (!this.modalBody) return;
        const zoomWrapper = this.modalBody.querySelector("[data-zoom-wrapper]") as HTMLDivElement;
        this.zoomLevel = Math.min(Math.max(z, 0.5), 3);
        zoomWrapper.style.transform = `scale(${this.zoomLevel})`;
    }

    /**
     * Revokes any previously created object URL for the current file.
     */
    private cleanupFileUrl(): void {
        if (this.currentFileUrl) {
            URL.revokeObjectURL(this.currentFileUrl);
            this.currentFileUrl = null;
        }
    }

    // -------------------------------------------------------------------------
    // Dataverse Web API helpers
    // -------------------------------------------------------------------------

    /**
     * Builds the Dataverse file download URL for the given record and file column.
     * Uses EntitySetName (configured in the manifest) to keep the control reusable.
     */
    private async getFileUrl(
        _entityLogicalName: string,
        entityId: string,
        fileColumnName: string
    ): Promise<string> {
        const entitySetName = this.context.parameters.EntitySetName.raw;
        const cleanId = (entityId || "").replace(/[{}]/g, "");

        const orgUrl =
            (this.context as any).page?.getClientUrl?.() ||
            (window as any).Xrm?.Utility?.getGlobalContext?.().getClientUrl?.() ||
            window.location.origin;

        const normalizedColumn = (fileColumnName || "").trim();

        const url = `${orgUrl}/api/data/v9.0/${entitySetName}(${cleanId})/${normalizedColumn}`;
        console.log("DocumentPreview – fileUrl:", url);
        return url;
    }

    /**
     * Retrieves the file from Dataverse using Web API metadata + file content endpoint.
     * Returns a Blob plus the file name and resolved MIME type.
     */
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

        // Dataverse $value endpoint can be used for direct download from the header button
        this.currentDownloadUrl = `${fileUrl}/$value`;

        // Retrieve Base64 content via Web API
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
    // File preview rendering (PDF / image / docx / fallback)
    // -------------------------------------------------------------------------

    /**
     * Renders the given file Blob into the modal body, using pdf.js, docx-preview
     * or a simple image element, with a clean, card-like layout.
     */
    private async renderFilePreview(file: { blob: Blob; fileName: string; mimeType: string }): Promise<void> {
        if (!this.modalBody) return;

        const zoomWrapper = this.modalBody.querySelector("[data-zoom-wrapper]") as HTMLDivElement;
        zoomWrapper.innerHTML = "";

        // Ensure base styling for inner container
        Object.assign(zoomWrapper.style, {
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            gap: "16px",
            padding: "16px",
            boxSizing: "border-box"
        });

        const lower = file.fileName.toLowerCase();
        const mime = file.mimeType.toLowerCase();
        console.log(file.fileName, file.mimeType, file.blob.size);

        // Helper to create a unified informational block (loading / error / fallback)
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
                textAlign: "center"
            });
            return div;
        };

        // ---------- PDF PREVIEW ----------
        if (mime === "application/pdf" || lower.endsWith(".pdf")) {
            const loading = createInfoBlock("Loading PDF preview...");
            zoomWrapper.appendChild(loading);

            try {
                const arrayBuffer = await file.blob.arrayBuffer();
                const pdf = await (pdfjsLib as any).getDocument({ data: arrayBuffer }).promise;

                zoomWrapper.innerHTML = "";

                const renderAllPages = true;
                const numPages = renderAllPages ? pdf.numPages : 1;

                for (let pageNum = 1; pageNum <= numPages; pageNum++) {
                    const page = await pdf.getPage(pageNum);

                    const viewport = page.getViewport({ scale: 1.2 });

                    const canvas = document.createElement("canvas");
                    const context2d = canvas.getContext("2d")!;
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    Object.assign(canvas.style, {
                        display: "block",
                        background: "#ffffff",
                        borderRadius: "4px"
                    });

                    // Wrap each page in a card-like container to match model-driven styling
                    const pageContainer = document.createElement("div");
                    Object.assign(pageContainer.style, {
                        backgroundColor: "#ffffff",
                        padding: "12px",
                        borderRadius: "6px",
                        boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                        margin: "0 auto",
                        display: "inline-block"
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

        // ---------- IMAGE PREVIEW ----------
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
                console.error("Error reading image blob for preview");
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

        // ---------- DOCX PREVIEW ----------
        if (lower.endsWith(".docx")) {
            const wordContainerOuter = document.createElement("div");
            Object.assign(wordContainerOuter.style, {
                backgroundColor: "#f5f5f5",
                padding: "12px",
                borderRadius: "6px",
                boxShadow: "0 1px 4px rgba(0,0,0,0.16)",
                maxWidth: "900px",
                width: "100%",
                boxSizing: "border-box"
            });

            const wordInner = document.createElement("div");
            Object.assign(wordInner.style, {
                backgroundColor: "#ffffff",
                padding: "16px",
                borderRadius: "4px",
                minHeight: "200px"
            });

            wordContainerOuter.appendChild(wordInner);
            zoomWrapper.appendChild(wordContainerOuter);

            await renderAsync(file.blob, wordInner);
            return;
        }

        // ---------- LEGACY .DOC PREVIEW ----------
        if (lower.endsWith(".doc")) {
            const msg = createInfoBlock(
                "This file is a legacy .doc format which cannot be rendered inline. Click the button below to open it in a new tab."
            );

            const btn = this.makeButton("Open document", () => {
                if (this.currentFileUrl) {
                    window.open(this.currentFileUrl, "_blank");
                }
            }, true);

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

        // ---------- FALLBACK ----------
        zoomWrapper.appendChild(
            createInfoBlock(`Preview is not available for this file type (${file.mimeType}).`)
        );
    }
}
