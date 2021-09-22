import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './BvMitteilungAlleWebPart.module.scss';
import * as strings from 'BvMitteilungAlleWebPartStrings';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  SPHttpClient, SPHttpClientConfiguration, SPHttpClientResponse,
  ODataVersion, ISPHttpClientConfiguration
} from '@microsoft/sp-http';

export interface ISPList {
  value: SPListItem[];
}
export interface SPListItem {
  FieldValuesAsText: { FileRef: string };
  Author: { Title: string };
  Title: string;
  Modified: string;
  File_x0020_Type: string;
  Created: string;
  Mitteilung: string;
  AttachmentFiles: { FileName: string, length: string };
  Id: string;
}

export interface IBvMitteilungAlleWebPartProps {
  description: string;
}

export default class BvMitteilungAlleWebPart extends BaseClientSideWebPart<IBvMitteilungAlleWebPartProps> {

  private SetButtonEventHandlers(): void {
    this.domElement.querySelectorAll("[id^='pkNewsItem_']").forEach((element) => {
      element.addEventListener('click', () => {
        let itemName = element.id.replace('pkNewsItem_', '');
        this.GetItem(itemName).then((response) => {
          response.value.forEach((item: SPListItem) => {
            let i = 0;
            let modified: string;
            let neuText: string = "Neu!";
            let files: string = `<div>`;

            let currentDate = new Date();
            let tempDate = new Date(item.Created);
            let endDate = new Date(tempDate.setDate(tempDate.getDate() + 7));

            if (item.AttachmentFiles.length != '1' && item.AttachmentFiles.length != '0') {
              files;
              for (i = 0; i < parseInt(item.AttachmentFiles.length); i++) {
                files += `<div style='margin-right: 20px'><a href='${this.context.pageContext.web.absoluteUrl}/Lists/BVMitteilungen/Attachments/${item.Id}/${item.AttachmentFiles[i].FileName}' target='_blank'>${item.AttachmentFiles[i].FileName}</a></div>`;
              }
            } else if (item.AttachmentFiles.length == '1') {
              files += `<div><a href='${this.context.pageContext.web.absoluteUrl}/Lists/BVMitteilungen/Attachments/${item.Id}/${item.AttachmentFiles[i].FileName}' target='_blank'>${item.AttachmentFiles[0].FileName}</a></div>`;
            } else { }

            files += `</div>`;

            if (currentDate > endDate) {
              neuText = '';
            }

            if (item.Mitteilung == null) {
              item.Mitteilung = '';
            }

            modified = item.Modified.substring(8, 10) + '-' + item.Modified.substring(5, 7) + '-' + item.Modified.substring(0, 4);

            let html = `<div>
              <div class='${styles.pkNewsResultHeader}'>
                <div><a id='pkNewsItem_${item.Title}'>${item.Title}</a></div>
                <input type='button' id='pkNewsDetailClose' class='${styles.pkNewsDetailClose}' value='X'/>
                <span>${neuText}</span>
              </div>
              <hr/>
              <div class='${styles.pkNewsResultCenter}'>
                <div style='word-break:break-all' class='pkNewsItemInhaltItem'>${item.Mitteilung}</div>
                <br>
                <div>von ${item.Author.Title}</div>
                <div>am ${modified}</div>
                <br>
                <div class='pkNewsAttachmentFilesItem'>${files}</div>
              </div>
            </div>`;

            const container: Element = this.domElement.querySelector('#pkNewsDetail') as HTMLElement;
            container.innerHTML += html;

            this.domElement.querySelectorAll("[class='pkNewsItemInhaltItem']").forEach((element) => {
              let links = element.getElementsByTagName('a');
              for (let i = 0; i < links.length; i++) {
                links[i].setAttribute('target', '_blank');
                if (links[i].href.toLowerCase().indexOf('.doc') !== -1 || links[i].href.toLowerCase().indexOf('.docx') !== -1
                  || links[i].href.toLowerCase().indexOf('.ppt') !== -1 || links[i].href.toLowerCase().indexOf('.pptx') !== -1
                  || links[i].href.toLowerCase().indexOf('.xls') !== -1 || links[i].href.toLowerCase().indexOf('.xlsx') !== -1
                  || links[i].href.toLowerCase().indexOf('.vsd') !== -1 || links[i].href.toLowerCase().indexOf('.vsdx') !== -1) {
                  links[i].href += '?web=1';
                }
              }
            });

            this.domElement.querySelectorAll("[class='pkNewsAttachmentFilesItem']").forEach((element) => {
              let links = element.getElementsByTagName('a');
              for (let i = 0; i < links.length; i++) {
                if (links[i].href.toLowerCase().indexOf('.doc') !== -1 || links[i].href.toLowerCase().indexOf('.docx') !== -1
                  || links[i].href.toLowerCase().indexOf('.ppt') !== -1 || links[i].href.toLowerCase().indexOf('.pptx') !== -1
                  || links[i].href.toLowerCase().indexOf('.xls') !== -1 || links[i].href.toLowerCase().indexOf('.xlsx') !== -1
                  || links[i].href.toLowerCase().indexOf('.vsd') !== -1 || links[i].href.toLowerCase().indexOf('.vsdx') !== -1) {
                  links[i].href += '?web=1';
                }
              }
            });
          });
        }).then(() => {
          let pkNewsDetail: HTMLElement = this.domElement.querySelector('#pkNewsDetail') as HTMLElement;
          pkNewsDetail.style.display = 'block';
          this.SetCloseButtonHandler();
        });
      });
    });
  }

  private SetCloseButtonHandler(): void {
    let pkNewsDetailClose = this.domElement.querySelector("[id^='pkNewsDetailClose']") as HTMLElement;
    try {
      let pkNewsDetailCloseNeu = pkNewsDetailClose.cloneNode(true);
      pkNewsDetailClose.parentNode.replaceChild(pkNewsDetailCloseNeu, pkNewsDetailClose);
      pkNewsDetailCloseNeu.addEventListener('click', () => {
        const pkNewsDetail: HTMLElement = this.domElement.querySelector('#pkNewsDetail') as HTMLElement;
        this.domElement.querySelector('#pkNewsDetail').innerHTML = '';
        pkNewsDetail.style.display = 'none';
      });
    } catch (error) { }
  }

  private RenderList(keyword: string): void {
    document.querySelector('#pkNewsAlleResult').innerHTML = '';
    this.GetListData(keyword).then((response) => {
      response.value.forEach((item: SPListItem) => {
        let i = 0;
        let modified: string;
        let neuText: string = "Neu!";
        let files: string = `<div>`;

        let currentDate = new Date();
        let tempDate = new Date(item.Created);
        let endDate = new Date(tempDate.setDate(tempDate.getDate() + 7));

        if (item.AttachmentFiles.length != '1' && item.AttachmentFiles.length != '0') {
          files;
          for (i = 0; i < parseInt(item.AttachmentFiles.length); i++) {
            files += `<div style='margin-right: 20px'><a href='${this.context.pageContext.web.absoluteUrl}/Lists/BVMitteilungen/Attachments/${item.Id}/${item.AttachmentFiles[i].FileName}' target='_blank'>${item.AttachmentFiles[i].FileName}</a></div>`;
          }
        } else if (item.AttachmentFiles.length == '1') {
          files += `<div><a href='${this.context.pageContext.web.absoluteUrl}/Lists/BVMitteilungen/Attachments/${item.Id}/${item.AttachmentFiles[i].FileName}' target='_blank'>${item.AttachmentFiles[0].FileName}</a></div>`;
        } else { }

        files += `</div>`;

        if (currentDate > endDate) {
          neuText = '';
        }

        if (item.Mitteilung == null) {
          item.Mitteilung = '';
        }

        modified = item.Modified.substring(8, 10) + '-' + item.Modified.substring(5, 7) + '-' + item.Modified.substring(0, 4);

        let html = `<div class='${styles.pkNewsResultCard}'>
          <div class='${styles.pkNewsResultHeader}'>
            <div><a id='pkNewsItem_${item.Id}' class='${styles.pkNewsItem}'>${item.Title}</a></div>
            <span>${neuText}</span>
          </div>
          <div class='${styles.pkNewsResultCenter}'>
            <div id='${styles.pkNewsItemInhalt}' class='pkNewsItemInhalt' style='word-break:break-all'>${item.Mitteilung}</div>
            <div>von ${item.Author.Title}</div>
            <div>am ${modified}</div>
            <div class='pkNewsAttachmentFiles'>${files}</div>
          </div>
        </div>`;

        const container: Element = this.domElement.querySelector('#pkNewsAlleResult') as HTMLElement;
        container.innerHTML += html;
      });
      this.domElement.querySelectorAll("[class='pkNewsItemInhalt']").forEach((element) => {
        let links = element.getElementsByTagName('a');
        for (let i = 0; i < links.length; i++) {
          links[i].setAttribute('target', '_blank');
          if (links[i].href.toLowerCase().indexOf('.doc') !== -1 || links[i].href.toLowerCase().indexOf('.docx') !== -1
            || links[i].href.toLowerCase().indexOf('.ppt') !== -1 || links[i].href.toLowerCase().indexOf('.pptx') !== -1
            || links[i].href.toLowerCase().indexOf('.xls') !== -1 || links[i].href.toLowerCase().indexOf('.xlsx') !== -1
            || links[i].href.toLowerCase().indexOf('.vsd') !== -1 || links[i].href.toLowerCase().indexOf('.vsdx') !== -1) {
            links[i].href += '?web=1';
          }
        }
      });

      this.domElement.querySelectorAll("[class='pkNewsAttachmentFiles']").forEach((element) => {
        let links = element.getElementsByTagName('a');
        for (let i = 0; i < links.length; i++) {
          if (links[i].href.toLowerCase().indexOf('.doc') !== -1 || links[i].href.toLowerCase().indexOf('.docx') !== -1
            || links[i].href.toLowerCase().indexOf('.ppt') !== -1 || links[i].href.toLowerCase().indexOf('.pptx') !== -1
            || links[i].href.toLowerCase().indexOf('.xls') !== -1 || links[i].href.toLowerCase().indexOf('.xlsx') !== -1
            || links[i].href.toLowerCase().indexOf('.vsd') !== -1 || links[i].href.toLowerCase().indexOf('.vsdx') !== -1) {
            links[i].href += '?web=1';
          }
        }
      });
    }).then(() => { this.SetButtonEventHandlers(); });
  }

  private async GetListData(keyword: string): Promise<ISPList> {
    if (keyword == '') {
      return this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('BV-Mitteilungen')/items?$orderby= Modified desc 
        &$select=Id,Title,Mitteilung,Created,Modified,Author/Title,AttachmentFiles/FileName,Attachments&$expand=Author,AttachmentFiles`,
        SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
    else {
      const spSearchConfig: ISPHttpClientConfiguration = {
        defaultODataVersion: ODataVersion.v3
      };
      const clientConfigODataV3: SPHttpClientConfiguration = SPHttpClient.configurations.v1.overrideWith(spSearchConfig);
      return this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('BV-Mitteilungen')/items?$select=Id,Title,Mitteilung,Created,Author/Title,Modified,AttachmentFiles
        &$expand=Author,AttachmentFiles&$orderby=Modified desc&$filter=substringof('${keyword}',Title) or substringof('${keyword}',Author/Title)`,
        clientConfigODataV3).then((response: SPHttpClientResponse) => {
          return response.json();
        });
    }
  }

  private GetItem(keyword: string): Promise<ISPList> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('BV-Mitteilungen')/items?$select=Id,Title,Mitteilung,Created,Author/Title,Modified,AttachmentFiles
      &$expand=Author,AttachmentFiles&$filter='${keyword}' eq Id`,
      SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.pkNewsAlle}">
        <div class='${styles.pkNewsDetail}' id='pkNewsDetail' style='display:none'></div>
        <div id='pkNewsAlleResult'></div>
      </div>`;

    if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
      this.RenderList('');
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
