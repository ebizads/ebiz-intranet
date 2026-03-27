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

    const hexCell = (
      icon: string,
      n: number,
      size: "sm" | "md" | "lg" = "lg",
      hidden = false,
    ): string => `
  <div class="${styles.hexCell} ${styles[size]} ${hidden ? styles.hexHidden : ""}">
    ${icon ? `<img src="${icon}" class="${styles.hexImage}" />` : ""}
    ${hexSvg(n)}
  </div>`;

    this.domElement.innerHTML = `
      <div class="${styles.heroContainer}">
        <h1 class="${styles.heroTitle}">
          We <span class="${styles.heroHighlight}">LISTEN, UNDERSTAND,<br/>COLLABORATE</span>
          to provide<br/>Innovative Solutions.
        </h1>

        <div class="${styles.hexGrid}">

          <div class="${styles.hexRow}  ${styles.hexRowShift}">
            ${hexCell("", 0, "lg", true)}
            ${hexCell("", 1, "lg", true)}
            ${hexCell("", 2, "lg")}
            ${hexCell("", 3, "lg")}
            ${hexCell("", 4, "lg", true)}
          </div>

          <div class="${styles.hexRow}">
            ${hexCell("", 5, "lg", true)}
            ${hexCell("", 6, "lg", true)}
            ${hexCell(bee, 7, "lg")}
            ${hexCell("", 8, "lg")}
            ${hexCell("", 9, "lg")}
          </div>

          <div class="${styles.hexRow}  ${styles.hexRowShift}">
            ${hexCell(honeyjar, 10, "lg")}
            ${hexCell("", 11, "lg")}
            ${hexCell(honey, 12, "lg")}
            ${hexCell("", 13, "lg")}
          </div>

          <div class="${styles.hexRow}">
            ${hexCell(honeyjar, 14, "lg")}
            ${hexCell("", 15, "lg")}
            ${hexCell("", 16, "lg")}
            ${hexCell("", 17, "lg")}
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
