import * as React from 'react';

import {
  useEffect,
  useState
} from 'react';

import { spfi, SPFI } from '@pnp/sp';

import { SPFx } from '@pnp/sp/presets/all';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import { IBankManagementProps } from './IBankManagementProps';

import { IBankItem } from './IBankItem';

const BankCustomer: React.FC<IBankManagementProps> = ({ context }) => {

  const [sp, setSp] = useState<SPFI>();

  const [customers, setCustomers] = useState<IBankItem[]>([]);

  const [editId, setEditId] = useState<number | null>(null);

  const [formData, setFormData] = useState<IBankItem>({

    CustomerName: '',
    Email: '',
    PhoneNumber: 0,
    AccountType: '',
    Balance: 0,
    IsActive: false,
    Branch: '',
    DateOfJoining: ''
  });

  useEffect(() => {

    const spInstance = spfi().using(SPFx(context));

    setSp(spInstance);

  }, [context]);

  useEffect(() => {

    if (sp) {

      getCustomers();
    }

  }, [sp]);

  const getCustomers = async (): Promise<void> => {

    try {

      const response: IBankItem[] = await sp!.web.lists
        .getByTitle('BankCustomers')
        .items();

      setCustomers(response);

    } catch (error) {

      console.log(error);
    }
  };

  const handleChange = (
    event: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>
  ): void => {

    const {
      name,
      value,
      type
    } = event.target;

    const checked = (event.target as HTMLInputElement).checked;

    setFormData({
      ...formData,
      [name]: type === 'checkbox' ? checked : value
    });
  };

  // EDIT FUNCTION

  const editCustomer = (item: IBankItem): void => {

    setEditId(item.Id!);

    setFormData({

      CustomerName: item.CustomerName,

      Email: item.Email,

      PhoneNumber: item.PhoneNumber,

      AccountType: item.AccountType,

      Balance: item.Balance,

      IsActive: item.IsActive,

      Branch: item.Branch,

      DateOfJoining: item.DateOfJoining
    });
  };

  // SAVE FUNCTION

  const saveData = async (): Promise<void> => {

    try {

      if (editId) {

        // UPDATE

        await sp!.web.lists
          .getByTitle('BankCustomers')
          .items
          .getById(editId)
          .update({

            CustomerName: formData.CustomerName,

            Email: formData.Email,

            PhoneNumber: formData.PhoneNumber,

            AccountType: formData.AccountType,

            Balance: formData.Balance,

            IsActive: formData.IsActive,

            Branch: formData.Branch,

            DateOfJoining: formData.DateOfJoining
          });

        alert('Customer Updated Successfully');

        setEditId(null);

      } else {

        // ADD

        await sp!.web.lists
          .getByTitle('BankCustomers')
          .items
          .add({

            CustomerName: formData.CustomerName,

            Email: formData.Email,

            PhoneNumber: formData.PhoneNumber,

            AccountType: formData.AccountType,

            Balance: formData.Balance,

            IsActive: formData.IsActive,

            Branch: formData.Branch,

            DateOfJoining: formData.DateOfJoining
          });

        alert('Customer Added Successfully');
      }

      // CLEAR FORM

      setFormData({

        CustomerName: '',
        Email: '',
        PhoneNumber: 0,
        AccountType: '',
        Balance: 0,
        IsActive: false,
        Branch: '',
        DateOfJoining: ''
      });

      // REFRESH TABLE

      getCustomers();

    } catch (error) {

      console.log(error);
    }
  };

  

  return (

    <div style={{ padding: '20px' }}>

      <h1>Bank Management</h1>

      <div style={{
        display: 'grid',
        gridTemplateColumns: '1fr 1fr',
        gap: '15px'
      }}>

        <input
          type='text'
          name='CustomerName'
          placeholder='Customer Name'
          value={formData.CustomerName}
          onChange={handleChange}
        />

        <input
          type='email'
          name='Email'
          placeholder='Email'
          value={formData.Email}
          onChange={handleChange}
        />

        <input
          type='number'
          name='PhoneNumber'
          placeholder='Phone Number'
          value={formData.PhoneNumber}
          onChange={handleChange}
        />

        <select
          name='AccountType'
          value={formData.AccountType}
          onChange={handleChange}
        >
          <option value=''>Select Account</option>

          <option value='Savings'>
            Savings
          </option>

          <option value='Current'>
            Current
          </option>

        </select>

        <input
          type='number'
          name='Balance'
          placeholder='Balance'
          value={formData.Balance}
          onChange={handleChange}
        />

        <select
          name='Branch'
          value={formData.Branch}
          onChange={handleChange}
        >
          <option value=''>Select Branch</option>

          <option value='Hyderabad'>
            Hyderabad
          </option>

          <option value='Chennai'>
            Chennai
          </option>

        </select>

        <input
          type='date'
          name='DateOfJoining'
          value={formData.DateOfJoining}
          onChange={handleChange}
        />

        <label>

          <input
            type='checkbox'
            name='IsActive'
            checked={formData.IsActive}
            onChange={handleChange}
          />

          Active

        </label>

      </div>

      <br />

      <button onClick={saveData}>

        {
          editId ? 'Update Customer' : 'Save Customer'
        }

      </button>
      <button onClick={() => setFormData({
        CustomerName: '',
        Email: '',
        PhoneNumber: 0,
        AccountType: '',
        Balance: 0,
        IsActive: false,
        Branch: '',
        DateOfJoining: ''
      })}>
        Clear
      </button>

      <hr />

      <table
        cellPadding={10}
        style={{ width: '100%' }}
      >

        <thead>

          <tr>

            <th>Customer Name</th>
            <th>Email</th>
            <th>Phone Number</th>
            <th>Account Type</th>
            <th>Balance</th>
            <th>Branch</th>
            <th>IsActive</th>
            <th>Actions</th>

          </tr>

        </thead>

        <tbody>

          {
            customers.map((item) => (

              <tr key={item.Id}>

                <td>{item.CustomerName}</td>

                <td>{item.Email}</td>

                <td>{item.PhoneNumber}</td>

                <td>{item.AccountType}</td>

                <td>{item.Balance}</td>
                <td>{item.Branch}</td>

                <td>{item.IsActive ? 'Yes' : 'No'}</td>

                <td>

                  <button
                    onClick={() => editCustomer(item)}
                  >
                    Edit
                  </button>

                </td>

              </tr>
            ))
          }

        </tbody>

      </table>

    </div>
  );
};
export default BankCustomer;