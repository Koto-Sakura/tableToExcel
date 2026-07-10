import {
    Plugin,
    showMessage,
} from "siyuan";
import "./index.scss";
import { copyTableBlockToClipboard } from "./clipboardExporter";
import { exportTableToExcel } from "./excelExporter";

export default class TableToExcelPlugin extends Plugin {
    private onTableBlockIconBindThis = this.onTableBlockIcon.bind(this);

    onload() {
        this.eventBus.on("click-blockicon", this.onTableBlockIconBindThis);
    }

    onunload() {
        this.eventBus.off("click-blockicon", this.onTableBlockIconBindThis);
    }

    private onTableBlockIcon({detail}: any) {
        const tableBlock = detail.blockElements.find(
            (el: HTMLElement) => el.dataset.type === "NodeTable",
        );
        if (!tableBlock) {
            return;
        }

        // 复制到剪贴板
        const copyItem = this.createMenuItem(
            "iconCopy",
            this.i18n.copyToExcel,
            () => {
                copyTableBlockToClipboard(tableBlock).then(() => {
                    window.siyuan.menus.menu.remove();
                    showMessage(this.i18n.copiedToExcel);
                }).catch((e: Error) => {
                    showMessage(`复制失败: ${e.message}`, 6000, "error");
                });
            },
        );
        detail.menu.addItem({ id: "tableToExcel_copy", element: copyItem });

        // 分隔线
        detail.menu.addSeparator();

        // 导出为 Excel 文件
        const exportItem = this.createMenuItem(
            "iconDownload",
            this.i18n.exportToExcel,
            () => {
                try {
                    exportTableToExcel(tableBlock);
                    window.siyuan.menus.menu.remove();
                } catch (e: any) {
                    showMessage(`导出失败: ${e.message}`, 6000, "error");
                }
            },
        );
        detail.menu.addItem({ id: "tableToExcel_export", element: exportItem });
    }

    private createMenuItem(iconName: string, label: string, onClick: () => void): HTMLDivElement {
        const el = document.createElement("div");
        el.className = "b3-menu__item";
        el.innerHTML = `<svg class="b3-menu__icon"><use xlink:href="#${iconName}"></use></svg><span class="b3-menu__label">${label}</span>`;
        el.addEventListener("click", onClick);
        return el;
    }
}
