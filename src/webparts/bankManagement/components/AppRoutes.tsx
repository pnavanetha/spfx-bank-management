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

        <Route
          path="/"
          element={<Home />}
        />


        <Route
          path="/customers"
          element={<BankCustomer {...props} />}
        />


        <Route
          path="/groups"
          element={<SharePointGroups {...props} />}
        />

      </Routes>

    </HashRouter>
  );
};

export default AppRoutes;