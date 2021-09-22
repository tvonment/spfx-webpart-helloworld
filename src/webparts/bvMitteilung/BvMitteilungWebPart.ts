import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './BvMitteilungWebPart.module.scss';
import * as strings from 'BvMitteilungWebPartStrings';

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

export interface IBvMitteilungWebPartProps {
  description: string;
}

import * as $ from 'jquery';

export default class BvMitteilungWebPart extends BaseClientSideWebPart<IBvMitteilungWebPartProps> {

  private SetButtonEventHandlers(): void {
    this.domElement.querySelector("[id^='pkNewsSearchButton']").addEventListener('click', () => {
      let keyword: string = $('#pkNewsSearchKeyword').val().toString();
      if (keyword == '') {
        keyword = null;
      }
      let pkNewsContent: HTMLElement = this.domElement.querySelector("[id^='pkNewsContent']") as HTMLElement;
      let pkNewsContentSearch: HTMLElement = this.domElement.querySelector("[id^='pkNewsContentSearch']") as HTMLElement;
      pkNewsContent.style.display = 'none';
      pkNewsContentSearch.style.display = 'block';

      const pkNewsContentSearchResult: HTMLElement = this.domElement.querySelector("[id^='pkNewsContentSearchResult']") as HTMLElement;
      pkNewsContentSearchResult.innerHTML = '';
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
              <div><a id='pkNewsItems_${item.Id}' class='${styles.pkNewsItem}'>${item.Title}</a></div>
              <span>${neuText}</span>
            </div>
            <div class='${styles.pkNewsResultCenter}'>
              <div id='${styles.pkNewsItemInhalt}' class='pkNewsItemInhaltSearchResult'>${item.Mitteilung}</div>
              <div>von ${item.Author.Title}</div>
              <div>am ${modified}</div>
              <div class='pkNewsAttachmentFilesSearch'>${files}</div>
            </div>
          </div>`;

          const container: Element = this.domElement.querySelector("[id^='pkNewsContentSearchResult']") as HTMLElement;
          container.innerHTML += html;
        });
        this.domElement.querySelectorAll("[class='pkNewsItemInhaltSearchResult']").forEach((element) => {
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

        this.domElement.querySelectorAll("[class='pkNewsAttachmentFilesSearch']").forEach((element) => {
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
      }).then(() => {
        this.SetSearchResultCloseButtonHandler();
        this.SetSearchResultButtonHandler();
      });
    });
    this.domElement.querySelector("[id^='pkNewsAlleButton']").addEventListener('click', () => {
      window.open(this.context.pageContext.web.absoluteUrl + `/SitePages/Alle-BV-Mitteilungen.aspx`);
    });
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
    }
    catch (error) { }
  }

  private SetSearchResultCloseButtonHandler(): void {
    let pkNewsCloseButton: HTMLElement = this.domElement.querySelector("[id^='pkNewsCloseButton']") as HTMLElement;
    try {
      let pkNewsCloseButtonNeu: HTMLElement = pkNewsCloseButton.cloneNode(true) as HTMLElement;
      pkNewsCloseButton.parentNode.replaceChild(pkNewsCloseButtonNeu, pkNewsCloseButton);
      pkNewsCloseButtonNeu.addEventListener('click', () => {
        const pkNewsContentSearch: HTMLElement = this.domElement.querySelector("[id^='pkNewsContentSearch']") as HTMLElement;
        pkNewsContentSearch.style.display = 'none';
        const pkNewsContent: HTMLElement = this.domElement.querySelector("[id^='pkNewsContent']") as HTMLElement;
        pkNewsContent.style.display = 'block';
        (this.domElement.querySelector("[id^='pkNewsSearchKeyword']") as HTMLInputElement).value = '';
      });
    } catch (error) { }
  }

  private SetSearchResultButtonHandler(): void {
    this.domElement.querySelectorAll("[id^='pkNewsItems_']").forEach((element) => {
      element.addEventListener('click', () => {
        let itemName = element.id.replace('pkNewsItems_', '');
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
                <div><a id='pkNewsItems_${item.Title}'>${item.Title}</a></div>
                <input type='button' id='pkNewsDetailClose' class='${styles.pkNewsDetailClose}' value='X'/>
                <span>${neuText}</span>
              </div>
              <hr/>
              <div class='${styles.pkNewsResultCenter}'>
                <div style='word-break:break-all' class='pkNewsItemInhaltSearch'>${item.Mitteilung}</div>
                <br>
                <div>von ${item.Author.Title}</div>
                <div>am ${modified}</div>
                <br>
                <div class='pkNewsAttachmentFilesSearchItem'>${files}</div>
              </div>
            </div>`;

            const container: Element = this.domElement.querySelector('#pkNewsDetail') as HTMLElement;
            container.innerHTML += html;
          });
          this.domElement.querySelectorAll("[class='pkNewsItemInhaltSearch']").forEach((element) => {
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

          this.domElement.querySelectorAll("[class='pkNewsAttachmentFilesSearchItem']").forEach((element) => {
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
        }).then(() => {
          let pkNewsDetail: HTMLElement = this.domElement.querySelector('#pkNewsDetail') as HTMLElement;
          pkNewsDetail.style.display = 'block';
          this.SetCloseButtonHandler();
        });
      });
    });
  }

  private RenderList(keyword: string): void {
    document.querySelector('#pkNewsContent').innerHTML = '';
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
            <div id='${styles.pkNewsItemInhalt}' class='pkNewsItemInhalt'>${item.Mitteilung}</div>
            <div>von ${item.Author.Title}</div>
            <div>am ${modified}</div>
            <div class='pkNewsAttachmentFiles'>${files}</div>
          </div>
        </div>`;
        const container: Element = this.domElement.querySelector('#pkNewsContent') as HTMLElement;
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
        &$select=Id,Title,Mitteilung,Created,Modified,Author/Title,AttachmentFiles/FileName&$expand=Author,AttachmentFiles&$top=2`,
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
      <div class="${styles.pkNews}">
        <div class='${styles.pkNewsTitle}'>
          <a href='${this.context.pageContext.web.absoluteUrl}/Lists/BVMitteilungen' target='_blank'>BV-Mitteilungen</a>
          <input class='${styles.pkNewsSearchKeyword}' id='pkNewsSearchKeyword' type='text' placeholder='Suche' />
          <input class='${styles.pkNewsSearchButton}' id='pkNewsSearchButton' type='button' value='Suchen' />
          <input class='${styles.pkNewsAlleButton}' id='pkNewsAlleButton' type='button' value='Alle ansehen' />
        </div>
        <hr/>
        <div class='${styles.pkNewsDetail}' id='pkNewsDetail' style='display:none'></div>
        <div class='${styles.pkNewsContent}' id='pkNewsContent'></div>
        <div class='${styles.pkNewsContentSearch}' id='pkNewsContentSearch'>
          Suchergebnisse: 
          <button id='pkNewsCloseButton' class='${styles.pkNewsCloseButton}'>X</button>
          <hr/>
          <div class='${styles.pkNewsContentSearchResult}' id='pkNewsContentSearchResult'></div>
        </div>
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
