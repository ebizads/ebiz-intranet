import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import bee from "./assets/bee.png";
import honeyjar from "./assets/honeyjar.png";
import honey from "./assets/honey.png";

import styles from "./DhiveheroWebPart.module.scss";
import * as strings from "DhiveheroWebPartStrings";

export interface IDhiveheroWebPartProps {
  description: string;
}

export default class DhiveheroWebPart extends BaseClientSideWebPart<IDhiveheroWebPartProps> {
  public render(): void {
    const uid = (this.instanceId || "x").replace(/-/g, "").substring(0, 8);

    const hexSvg = (n: number): string => `
      <svg class="${styles.hexSvg}" viewBox="0 0 100 86" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <linearGradient id="hg${uid}${n}" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#47A247"/>
            <stop offset="100%" stop-color="#EE7623"/>
          </linearGradient>
          <linearGradient id="hl${uid}${n}" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="rgba(255,255,255,0.45)"/>
            <stop offset="100%" stop-color="rgba(255,255,255,0)"/>
          </linearGradient>
        </defs>
        <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
          fill="#f5f5f5"
          stroke="url(#hg${uid}${n})"
          stroke-width="2"/>
        <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
          fill="url(#hl${uid}${n})"/>
      </svg>`;

    const hexCell = (icon: string, n: number): string => `
      <div class="${styles.hexCell}">
        ${icon ? `<img src="${icon}" class="${styles.hexImage}" />` : ""}
        ${hexSvg(n)}
      </div>`;

    this.domElement.innerHTML = `
      <div class="${styles.heroContainer}">
        <h1 class="${styles.heroTitle}">
          We <span class="${styles.heroHighlight}">LISTEN, UNDERSTAND,<br/>COLLABORATE</span>
          to provide<br/>Innovative Solutions.
        </h1>
  <a href="#" class="${styles.hexLink}" style="top:10%; right:35%; width:180px; height:155px;">
    <img src="${honeyjar}"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#7fb069"/>
    </svg>
  </a>
        <div class="${styles.hexGrid}">

          <div class="${styles.hexRow}">
            ${hexCell("", 1)}
            ${hexCell(bee, 2)}
            ${hexCell("", 3)}
            ${hexCell("", 4)}
          </div>

          <div class="${styles.hexRow} ${styles.hexRowShift}">
            ${hexCell(honeyjar, 5)}
            ${hexCell("", 6)}
            ${hexCell(honey, 7)}
            ${hexCell("", 8)}
          </div>

          <div class="${styles.hexRow}">
            ${hexCell("", 9)}
            ${hexCell(honeyjar, 10)}
            ${hexCell("", 11)}
            ${hexCell("", 12)}
          </div>

          <div class="${styles.hexRow} ${styles.hexRowShift}">
            ${hexCell(bee, 13)}
            ${hexCell("", 14)}
            ${hexCell("", 15)}
          </div>

        </div>
      </div>
    `;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
