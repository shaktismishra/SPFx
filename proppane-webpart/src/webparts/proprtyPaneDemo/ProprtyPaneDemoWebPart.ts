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

import styles from './ProprtyPaneDemoWebPart.module.scss';
import * as strings from 'ProprtyPaneDemoWebPartStrings';

export interface IProprtyPaneDemoWebPartProps {
  description: string;
  multiline: string;
  checkbox: string;
  dropdown: string;
  toggle: string;
}

export default class ProprtyPaneDemoWebPart extends BaseClientSideWebPart<IProprtyPaneDemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.proprtyPaneDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">Description: ${escape(this.properties.description)}</p>
              <p class="${ styles.description }">Multiline: ${escape(this.properties.multiline)}</p>
              <p class="${ styles.description }">Checkbox: ${escape(this.properties.checkbox)}</p>
              <p class="${ styles.description }">Dropdown: ${escape(this.properties.dropdown)}</p>
              <p class="${ styles.description }">Toggle: ${escape(this.properties.toggle)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
              PropertyPaneTextField('multiline', {
                label: 'Multi-line Text Field',
                multiline: true
              }),
              PropertyPaneCheckbox('checkbox', {
                text: 'Checkbox'
              }),
              PropertyPaneDropdown('dropdown', {
                label: 'Dropdown',
                options: [
                  { key: '1', text: 'One' },
                  { key: '2', text: 'Two' },
                  { key: '3', text: 'Three' },
                  { key: '4', text: 'Four' }
                ]}),
              PropertyPaneToggle('toggle', {
                label: 'Toggle',
                onText: 'On',
                offText: 'Off'
              })
            ]
            }
          ]
        }
      ]
    };
  }
}
