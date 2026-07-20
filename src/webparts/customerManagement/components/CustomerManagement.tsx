import * as React from 'react';
import { useEffect, useState } from 'react';

import type { ICustomerManagementProps } from './ICustomerManagementProps';
import SharePointService from '../services/SharePointService';
import { ICustomer } from '../models/ICustomer';

const CustomerManagement: React.FC<ICustomerManagementProps> = (props) => {

  const [customers, setCustomers] = useState<ICustomer[]>([]);

  // Add Customer States
  const [showForm, setShowForm] = useState(false);

  const [title, setTitle] = useState("");
  const [address, setAddress] = useState("");
  const [numberOfRigs, setNumberOfRigs] = useState(0);
  const [numberOfJackUps, setNumberOfJackUps] = useState(0);
  const [numberOfMODUs, setNumberOfMODUs] = useState(0);
  const [websiteURL, setWebsiteURL] = useState("");

  const [isEdit, setIsEdit] = useState(false);
  const [editId, setEditId] = useState<number>(0);

  const resetForm = (): void => {
    setIsEdit(false);
    setEditId(0);

    setTitle("");
    setAddress("");
    setNumberOfRigs(0);
    setNumberOfJackUps(0);
    setNumberOfMODUs(0);
    setWebsiteURL("");
  };

  const loadCustomers = async (): Promise<void> => {

    try {

      const service = new SharePointService(
        props.siteUrl,
        props.spHttpClient
      );

      const result = await service.getCustomers();

      setCustomers(result);

    } catch (error) {

      console.error("Error loading customers:", error);

    }

  };

  const updateCustomer = async (): Promise<void> => {

    try {

      const service = new SharePointService(
        props.siteUrl,
        props.spHttpClient
      );

      await service.updateCustomer(editId, {
        Title: title,
        Address: address,
        NumberOfRigs: numberOfRigs,
        NumberOfJackUps: numberOfJackUps,
        NumberOfMODUs: numberOfMODUs,
        WebsiteURL: websiteURL
      });

      alert("Customer updated successfully.");

      setShowForm(false);
      resetForm();

      await loadCustomers();

    } catch (error) {

      console.error(error);

    }

  };

  const deleteCustomer = async (id: number): Promise<void> => {

    const confirmDelete = window.confirm(
      "Are you sure you want to delete this customer?"
    );

    if (!confirmDelete) {
      return;
    }

    try {

      const service = new SharePointService(
        props.siteUrl,
        props.spHttpClient
      );

      await service.deleteCustomer(id);

      alert("Customer deleted successfully.");

      await loadCustomers();

    } catch (error) {

      console.error(error);

    }

  };

  // ADD CUSTOMER
  const addCustomer = async (): Promise<void> => {

    try {

      const service = new SharePointService(
        props.siteUrl,
        props.spHttpClient
      );

      await service.addCustomer({
        Title: title,
        Address: address,
        NumberOfRigs: numberOfRigs,
        NumberOfJackUps: numberOfJackUps,
        NumberOfMODUs: numberOfMODUs,
        WebsiteURL: websiteURL
      });

      alert("Customer added successfully.");

      resetForm();
      setShowForm(false);

      await loadCustomers();

    } catch (error) {

      console.error("Error adding customer:", error);

    }

  };

  const editCustomer = (customer: ICustomer): void => {

    setIsEdit(true);
    setEditId(customer.Id!);

    setTitle(customer.Title);
    setAddress(customer.Address);
    setNumberOfRigs(customer.NumberOfRigs);
    setNumberOfJackUps(customer.NumberOfJackUps);
    setNumberOfMODUs(customer.NumberOfMODUs);
    setWebsiteURL(customer.WebsiteURL);

    setShowForm(true);

  };

  useEffect(() => {

    loadCustomers().catch((error) => {
      console.error(error);
    });

  }, []);

  const tableStyle: React.CSSProperties = {
    borderCollapse: "collapse",
    width: "100%",
    border: "1px solid #ccc",
    marginTop: "20px"
  };

  const cellStyle: React.CSSProperties = {
    border: "1px solid #ccc",
    padding: "10px",
    textAlign: "left"
  };

  return (

    <div style={{ padding: "20px" }}>

      <h1>Customer Management</h1>

      <h3>Fennex Onboarding Project</h3>

      <p>Welcome {props.userDisplayName} 👋</p>

      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "20px"
        }}
      >
        <button
          onClick={() => {
            resetForm();
            setShowForm(true);
          }}
          style={{
            padding: "8px 15px",
            backgroundColor: "#0078d4",
            color: "white",
            border: "none",
            cursor: "pointer",
            borderRadius: "4px"
          }}
        >
          Add Customer
        </button>

        <div>

          <button
            style={{
              marginRight: "10px",
              padding: "8px 15px",
              backgroundColor: "green",
              color: "white",
              border: "none",
              cursor: "pointer",
              borderRadius: "4px"
            }}
            onClick={() =>
              window.open("https://localhost:7061/api/Export/excel", "_blank")
            }
          >
            Export Excel
          </button>

          <button
            style={{
              padding: "8px 15px",
              backgroundColor: "#d83b01",
              color: "white",
              border: "none",
              cursor: "pointer",
              borderRadius: "4px"
            }}
            onClick={() =>
              window.open("https://localhost:7061/api/Export/word", "_blank")
            }
          >
            Export Word
          </button>

        </div>

      </div>

      {showForm && (

        <div
          style={{
            border: "1px solid #ccc",
            padding: "20px",
            marginTop: "20px",
            width: "400px"
          }}
        >

          <h3>{isEdit ? "Edit Customer" : "Add Customer"}</h3>

          <div>
            <label>Customer Name</label>
            <br />
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
            />
          </div>

          <br />

          <div>
            <label>Address</label>
            <br />
            <input
              type="text"
              value={address}
              onChange={(e) => setAddress(e.target.value)}
            />
          </div>

          <br />

          <div>
            <label>Number Of Rigs</label>
            <br />
            <input
              type="number"
              value={numberOfRigs}
              onChange={(e) => setNumberOfRigs(Number(e.target.value))}
            />
          </div>

          <br />

          <div>
            <label>Number Of JackUps</label>
            <br />
            <input
              type="number"
              value={numberOfJackUps}
              onChange={(e) => setNumberOfJackUps(Number(e.target.value))}
            />
          </div>

          <br />

          <div>
            <label>Number Of MODUs</label>
            <br />
            <input
              type="number"
              value={numberOfMODUs}
              onChange={(e) => setNumberOfMODUs(Number(e.target.value))}
            />
          </div>

          <br />

          <div>
            <label>Website URL</label>
            <br />
            <input
              type="text"
              value={websiteURL}
              onChange={(e) => setWebsiteURL(e.target.value)}
            />
          </div>

          <br />

          <button
            onClick={() => {

              if (isEdit) {

                updateCustomer().catch((error) => {
                  console.error(error);
                });

              } else {

                addCustomer().catch((error) => {
                  console.error(error);
                });

              }

            }}
          >
            {isEdit ? "Update" : "Save"}
          </button>

          <button
            style={{ marginLeft: "10px" }}
            onClick={() => {
              resetForm();
              setShowForm(false);
            }}
          >
            Cancel
          </button>

        </div>

      )}

      <table style={tableStyle}>

        <thead>
          <tr>
            <th style={cellStyle}>Customer</th>
            <th style={cellStyle}>Address</th>
            <th style={cellStyle}>Number Of Rigs</th>
            <th style={cellStyle}>Number Of JackUps</th>
            <th style={cellStyle}>Number Of MODUs</th>
            <th style={cellStyle}>Website</th>
            <th style={cellStyle}>Actions</th>
          </tr>
        </thead>

        <tbody>

          {customers.length > 0 ? (

            customers.map((customer) => (

              <tr key={customer.Id}>

                <td style={cellStyle}>{customer.Title}</td>
                <td style={cellStyle}>{customer.Address}</td>
                <td style={cellStyle}>{customer.NumberOfRigs}</td>
                <td style={cellStyle}>{customer.NumberOfJackUps}</td>
                <td style={cellStyle}>{customer.NumberOfMODUs}</td>

                <td style={cellStyle}>
                  <a
                    href={customer.WebsiteURL}
                    target="_blank"
                    rel="noreferrer"
                  >
                    {customer.WebsiteURL}
                  </a>
                </td>

                <td style={cellStyle}>

                  <button
                    onClick={() => editCustomer(customer)}
                  >
                    Edit
                  </button>

                  <button
                    style={{ marginLeft: "10px" }}
                    onClick={() => {
                      deleteCustomer(customer.Id!).catch((error) => {
                        console.error(error);
                      });
                    }}
                  >
                    Delete
                  </button>

                </td>

              </tr>

            ))

          ) : (

            <tr>
              <td
                style={cellStyle}
                colSpan={7}
              >
                No customers found.
              </td>
            </tr>

          )}

        </tbody>

      </table>

    </div>

  );

};

export default CustomerManagement;