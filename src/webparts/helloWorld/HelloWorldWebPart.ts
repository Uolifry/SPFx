import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorld.module.scss';
import * as strings from 'helloWorldStrings';
import { IHelloWorldWebPartProps } from './IHelloWorldWebPartProps';
import MockHttpClient from './MockHttpClient';
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.helloWorld}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
              <p class="ms-font-l ms-fontColor-white">Loading from ${escape(this.context.pageContext.web.title)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>  
        <div>
          <div>Fibonachi <3 </div>
          <div>Fib 0 => ${this._countFibonachFlat(0).join(", ")}</div>
          <div>Fib 1 => ${this._countFibonachFlat(1).join(", ")}</div>
          <div>Fib 10 => ${this._countFibonachFlat(10).join(", ")}</div>
          <div>Fib N => ${this._fibonachRecursiveExecuter(parseInt(this.properties.fibonachi)).join(", ")}</div>
          <div>Fibonachi <3 recusive </div>
          <div>Fib 0 => ${this._fibonachRecursiveExecuter(0).join(", ")}</div>
          <div>Fib 1 => ${this._fibonachRecursiveExecuter(1).join(", ")}</div>
          <div>Fib 10 => ${this._fibonachRecursiveExecuter(10).join(", ")}</div>
          <div>Fib N => ${this._fibonachRecursiveExecuter(parseInt(this.properties.fibonachi)).join(", ")}</div>
        </div>
        <div id="spListContainer" />
      </div>`;

    this._renderListAsync();
  }

  private _countFibonachFlat(size: number) : number[] {
      let number0 = 0;
      let number1 = 1;
      let result  = [];
      let counter  = 0;
      let numberN = null;
      while(counter <= size ) {
        if(counter==0){
          result.push(number0);
        }else{
          if(counter == 1){
            result.push(number1);
          }else{
            numberN = number0 + number1;
            result.push(numberN);
            number0 = number1;
            number1 = numberN;
          }
        }
        counter++;
      }
      return result;
  }

  private _fibonachRecursiveExecuter(size:number): number[] {
    let storage = [];
    try {
      this._countFibonachRecusive(size, storage);  
    } catch (error) {
      storage = [error];
    }
    return storage;
 }
   
  private _countFibonachRecusive(size : number, storage: number[] ): number {
      let value;  
      if(size <= 1){
        if(size == 0)  { 
          value = 0;
        }
        if(size == 1 ){
          value = 1;
          this._countFibonachRecusive(size-1, storage);
        }
      }else{
        value = this._countFibonachRecusive(size-1, storage) + this._countFibonachRecusive(size-2, storage);
      }
      storage[size] = value;
      return value;
  }

  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse ) => {
        return response.json();
      });
  }

  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    else if (Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
		<ul class="${styles.list}">
			<li class="${styles.listItem}">
				<span class="ms-font-l">${item.Title}</span>
			</li>
		</ul>`;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                  label: 'Description'
                }),
                PropertyPaneTextField('test', {
                  label: 'Multi-line Text Field',
                  multiline: true
                }),
                PropertyPaneCheckbox('test1', {
                  text: 'Checkbox'
                }),
                PropertyPaneDropdown('test2', {
                  label: 'Dropdown',
                  options: [
                    { key: '1', text: 'One' },
                    { key: '2', text: 'Two' },
                    { key: '3', text: 'Three' },
                    { key: '4', text: 'Four' }
                  ]
                }),
                PropertyPaneToggle('test3', {
                  label: 'Toggle',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('fibonachi',{
                  label : 'Fibonachi N'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}