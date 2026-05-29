import * as React from 'react';

import { useNavigate } from 'react-router-dom';

const Home: React.FC = () => {

  const navigate = useNavigate();

  return (

    <div
      style={{
        padding: '40px',
        textAlign: 'center'
      }}
    >

      <h1>🏦 Bank Management System</h1>

      <p>Welcome to Home Page</p>

      <button
        onClick={() => navigate('/customers')}
        style={{
          padding: '12px 25px',
          cursor: 'pointer',
          fontSize: '16px'
        }}
      >
        Go To Customers
      </button>
      <button onClick={() => navigate('/groups')}
        style={{
          padding: '12px 25px',
          cursor: 'pointer',
          fontSize: '16px'
        }}
      >
        SharePoint Groups
      </button>
      <button onClick={() => navigate('/customer-masters')}
        style={{
          padding: '12px 25px',
          cursor: 'pointer',
          fontSize: '16px'
        }}
       >
        Customer Masters
      </button>

    </div>
  );
};

export default Home;