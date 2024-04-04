import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './CrmChangeRequestWebPart.module.scss';
import * as strings from 'CrmChangeRequestWebPartStrings';

export interface ICrmChangeRequestWebPartProps {
  description: string;
}

export default class CrmChangeRequestWebPart extends BaseClientSideWebPart<ICrmChangeRequestWebPartProps> {


  private _companyClients: { id: number, name: string }[] = [];
  private _companyProjects: { id: number, name: string }[] = [];
  private _companyTasks: string[] = [];

  private _fetchCompanyClients(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ClientName')/items?$select=ID,Client`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            throw new Error('Failed to fetch Company Clients');
          }
        })
        .then((data: any) => {
          this._companyClients = data.value.map((item: any) => ({ id: item.ID, name: item.Client }));
          this.renderDropdown(); // Ensure rendering after data is fetched
          resolve();
        })
        .catch((error: any) => {
          console.error('Error fetching Company Clients:', error);
          reject(error);
        });
    });
  }

  private _fetchCompanyProjects(client: string): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ClientProject')/items?$select=ID,Project&$filter=Client eq '${client}'`;
      this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            throw new Error('Failed to fetch Company Projects');
          }
        })
        .then((data: any) => {
          this._companyProjects = data.value.map((item: any) => ({ id: item.ID, name: item.Project }));
          resolve();
        })
        .catch((error: any) => {
          console.error('Error fetching Company Projects:', error);
          reject(error);
        });
    });
  }

  private _fetchCompanyTasks(projectId: number): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ProjectTask')/items?$select=Task&$filter=ProjectId eq '${projectId}'`;
      this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            throw new Error('Failed to fetch Company Tasks');
          }
        })
        .then((data: any) => {
          this._companyTasks = data.value.map((item: any) => item.Task);
          resolve();
        })
        .catch((error: any) => {
          console.error('Error fetching Company Tasks:', error);
          reject(error);
        });
    });
  }

  private renderDropdown(): void {
    const dropdownContent = this.domElement.querySelector(`.${styles.dropdownContent}`) as HTMLElement;
    if (dropdownContent) {
      // Render the "Client" dropdown here
    }
  }

  public render(): void {
    this._fetchCompanyClients().then(() => {
      this.domElement.innerHTML = `
      <div class="${styles.container}">
        <h2>Task Submission Form</h2>     
        <form id="taskForm">
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="client">Client:</label>
            <select class="${styles.formGroupinput}" id="client" name="client">
            ${this._companyClients.map(client => `<option class="${styles.dropdownItem}" value="${client.id}">${client.name}</option>`).join('')}
            </select>
          </div>
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="project">Project:</label>
            <select class="${styles.formGroupinput}" id="project" name="project"></select>
          </div>
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="task">Task:</label>
            <select class="${styles.formGroupinput}" id="task" name="task"></select>
          </div>
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="requestType">Request Type :</label>
            <input class="${styles.formGroupinput}" type="text" id="requestType" name="requestType">
          </div>
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="startDate">Start Date:</label>
            <input class="${styles.formGroupinput}" type="date" id="startDate" name="startDate"></input>
          </div>
          <div class="${styles.formGroup}">
            <label class="${styles.formGrouplabel}" for="description">Description:</label>
            <textarea class="${styles.formGroupinput}" type="text" id="description" name="description" rows="4"></textarea>
          </div>
                  
          <button  class="${styles.btn}" type="submit">Submit</button>
          
        </form>
      </div>`;

      const clientDropdown = this.domElement.querySelector('#client') as HTMLSelectElement;
      if (clientDropdown) {
        clientDropdown.addEventListener('change', () => {
          const selectedClient = clientDropdown.value;
          this._fetchCompanyProjects(selectedClient).then(() => {
            this.renderProjectDropdown();
          });
        });
      }

      const projectDropdown = this.domElement.querySelector('#project') as HTMLSelectElement;
      if (projectDropdown) {
        projectDropdown.addEventListener('change', () => {
          const selectedProjectId = parseInt(projectDropdown.value);
          this._fetchCompanyTasks(selectedProjectId).then(() => {
            this.renderTaskDropdown();
          });
        });
      }

      this.setupFormSubmitHandler();
    });
  }
  

  private renderProjectDropdown(): void {
    const projectDropdown = this.domElement.querySelector('#project') as HTMLSelectElement;
    if (projectDropdown) {
      projectDropdown.innerHTML = ''; // Clear previous options
      this._companyProjects.forEach(project => {
        const option = document.createElement('option');
        option.value = project.id.toString();
        option.text = project.name;
        projectDropdown.appendChild(option);
      });
    }
  }

  private renderTaskDropdown(): void {
    const taskDropdown = this.domElement.querySelector('#task') as HTMLSelectElement;
    if (taskDropdown) {
      taskDropdown.innerHTML = ''; // Clear previous options
      this._companyTasks.forEach(task => {
        const option = document.createElement('option');
        option.value = task;
        option.text = task;
        taskDropdown.appendChild(option);
      });
    }
  }


  private setupFormSubmitHandler(): void {
    const form = this.domElement.querySelector('#taskForm') as HTMLFormElement;

    if (form) {
      const self = this; // Store a reference to 'this'

      form.addEventListener('submit', async function(event) {
        event.preventDefault();
        
        const client = (form.querySelector('#client') as HTMLInputElement).value;
        const project = (form.querySelector('#project') as HTMLInputElement).value;
        const task = (form.querySelector('#task') as HTMLInputElement).value;
        const requestType = (form.querySelector('#requestType') as HTMLInputElement).value;
        const startDate = (form.querySelector('#startDate') as HTMLInputElement).value;
        const description = (form.querySelector('#description') as HTMLTextAreaElement).value;

        // Call the method to submit data
        await self.makePOSTRequest(client, project, task, requestType, startDate, description);

        // Reset form fields on successful submission
        form.reset();
      });
    } else {
      console.error('Form element not found.');
    }
}

private async makePOSTRequest(client: string, project: string, task: string, requestType: string, startDate: string, description: string): Promise<void> {
  try {
    const taskId = await this.getLookupItemId('ProjectTask', 'Task', task);

    if (!taskId) {
      console.error('Error: Unable to find Task ID for the given task.');
      return;
    }

    const requestBody = {
      Description: description, // Mapping 'Description' to the appropriate column
      ClientId: client,
      ProjectId: project,
      TaskId: taskId,
      RequestType: requestType,
      StartDate: startDate
    };

    const requestOptions: ISPHttpClientOptions = {
      body: JSON.stringify(requestBody),
      headers: {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata'
      }
    };

    const response = await this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CRMList')/items`, SPHttpClient.configurations.v1, requestOptions
    );

    if (response.ok) {
      console.log('Data submitted successfully.');
    } else {
      console.error('Failed to submit data. Status:', response.status, 'Message:', response.statusText);
    }
  } catch (error) {
    console.error('Error submitting data:', error);
  }
}




private async getLookupItemId(listName: string, lookupField: string, lookupValue: string): Promise<number | null> {
  try {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Id&$filter=${lookupField} eq '${lookupValue}'`,
      SPHttpClient.configurations.v1
    );

    if (response.ok) {
      const data = await response.json();
      if (data && data.value && data.value.length > 0) {
        return data.value[0].Id;
      }
    } else {
      console.error(`Failed to fetch ID for ${lookupValue}. Status: ${response.status}, Message: ${response.statusText}`);
    }
  } catch (error) {
    console.error('Error fetching lookup item ID:', error);
  }

  return null;
}

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
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
