import * as React from 'react';

import {
  HashRouter,
  Routes,
  Route
} from 'react-router-dom';

import Home from './Home';
import BankCustomer from './BankCustomer';
import SharePointGroups from './SharePointGroups';

import { IBankManagementProps } from './IBankManagementProps';

const AppRoutes: React.FC<IBankManagementProps> = (props) => {

  return (

    <HashRouter>

      <Routes>

        {/* HOME PAGE */}

        <Route
          path="/"
          element={<Home />}
        />

        {/* CUSTOMER PAGE */}

        <Route
          path="/customers"
          element={<BankCustomer {...props} />}
        />

        {/* SHAREPOINT GROUPS */}

        <Route
          path="/groups"
          element={<SharePointGroups {...props} />}
        />

      </Routes>

    </HashRouter>
  );
};

export default AppRoutes;