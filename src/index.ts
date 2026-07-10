import {
    Plugin,
    showMessage,
} from "siyuan";
import "./index.scss";
import { copyTableBlockToClipboard } from "./clipboardExporter";

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

        const itemElement = document.createElement("div");
        itemElement.className = "b3-menu__item";
        itemElement.innerHTML = `<svg class="b3-menu__icon"><use xlink:href="#iconCopy"></use></svg><span class="b3-menu__label">${this.i18n.copyToExcel}</span>`;
        itemElement.addEventListener("click", () => {
            copyTableBlockToClipboard(tableBlock).then(() => {
                window.siyuan.menus.menu.remove();
                showMessage(this.i18n.copiedToExcel);
            }).catch((e: Error) => {
                showMessage(`复制失败: ${e.message}`, 6000, "error");
            });
        });

        detail.menu.addItem({
            id: "tableToExcel_copy",
            element: itemElement,
        });
    }
}
