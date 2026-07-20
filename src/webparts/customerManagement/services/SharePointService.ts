import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICustomer } from "../models/ICustomer";

export default class SharePointService {

  constructor(
    private siteUrl: string,
    private spHttpClient: SPHttpClient
  ) { }

  // Get Customers
  public async getCustomers(): Promise<ICustomer[]> {

    const url =
      `${this.siteUrl}/_api/web/lists/GetByTitle('CustomerInformation')/items`;

    const response: SPHttpClientResponse =
      await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

    if (!response.ok) {
      throw new Error("Unable to fetch customers");
    }

    const result = await response.json();

    console.log("Customers:", result.value);

    return result.value;
  }

  // Add Customer
  public async addCustomer(customer: ICustomer): Promise<void> {

    const url =
      `${this.siteUrl}/_api/web/lists/GetByTitle('CustomerInformation')/items`;

    const response: SPHttpClientResponse =
      await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata"
          },
          body: JSON.stringify({
            Title: customer.Title,
            Address: customer.Address,
            NumberOfRigs: customer.NumberOfRigs,
            NumberOfJackUps: customer.NumberOfJackUps,
            NumberOfMODUs: customer.NumberOfMODUs,
            WebsiteURL: customer.WebsiteURL
          })
        }
      );

    if (!response.ok) {
      throw new Error("Unable to add customer");
    }

    console.log("Customer added successfully.");
  }

  // Update Customer
  public async updateCustomer(
    id: number,
    customer: ICustomer
  ): Promise<void> {

    const url =
      `${this.siteUrl}/_api/web/lists/GetByTitle('CustomerInformation')/items(${id})`;

    const response: SPHttpClientResponse =
      await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
          },
          body: JSON.stringify({
            Title: customer.Title,
            Address: customer.Address,
            NumberOfRigs: customer.NumberOfRigs,
            NumberOfJackUps: customer.NumberOfJackUps,
            NumberOfMODUs: customer.NumberOfMODUs,
            WebsiteURL: customer.WebsiteURL
          })
        }
      );

    if (!response.ok) {
      throw new Error("Unable to update customer");
    }

    console.log("Customer updated successfully.");
  }

  // Delete Customer
  public async deleteCustomer(id: number): Promise<void> {

    const url =
      `${this.siteUrl}/_api/web/lists/GetByTitle('CustomerInformation')/items(${id})`;

    const response: SPHttpClientResponse =
      await this.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          }
        }
      );

    if (!response.ok) {
      throw new Error("Unable to delete customer");
    }

    console.log("Customer deleted successfully.");
  }
}