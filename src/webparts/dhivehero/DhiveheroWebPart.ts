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
    this.domElement.innerHTML = `
    <div class="${styles.heroSection} ">
      
      <div class="${styles.heroText}">
        <h1>
          We <span class="${styles.gradientHeading}"> LISTEN, UNDERSTAND, COLLABORATE</span>
          to provide Innovative Solutions.
        </h1>
      </div>

<div class="${styles.hexWrapper}">

  <!-- HEX 12 -->
  <a href="#" class="${styles.hexLink}" style="top:18%; right:8%; width:280px; height:240px;">
    <img src="${bee}"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#f28c28"/>
    </svg>
  </a>

  <!-- HEX 2 -->
  <a href="#" class="${styles.hexLink}" style="top:10%; right:35%; width:180px; height:155px;">
    <img src="${honeyjar}"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#7fb069"/>
    </svg>
  </a>

  <!-- HEX 3 -->
  <a href="#" class="${styles.hexLink}" style="top:5%; right:20%; width:100px; height:86px;">
    <img src="${honey}"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#f28c28"/>
    </svg>
  </a>

  <!-- HEX 4 -->
  <a href="#" class="${styles.hexLink}" style="bottom:15%; right:30%; width:240px; height:206px;">
    <img src="https://via.placeholder.com/400"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#7fb069"/>
    </svg>
  </a>

  <!-- HEX 5 -->
  <a href="#" class="${styles.hexLink}" style="bottom:5%; right:12%; width:160px; height:138px;">
    <img src="https://via.placeholder.com/300"
         class="${styles.hexImage}" />
    <svg class="${styles.hexSvg}" viewBox="0 0 100 86">
      <polygon points="25,1 75,1 99,43 75,85 25,85 1,43"
        class="${styles.hexBorder}"
        stroke="#f28c28"/>
    </svg>
  </a>

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
