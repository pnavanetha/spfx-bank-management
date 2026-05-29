import * as React from 'react';

import { useEffect, useState } from 'react';

import { spfi, SPFI } from '@pnp/sp';

import { SPFx } from '@pnp/sp/presets/all';

import '@pnp/sp/webs';
import '@pnp/sp/site-groups';

import { useNavigate } from 'react-router-dom';

import { IBankManagementProps } from './IBankManagementProps';
// import { group } from 'console';

interface IGroup {
  Id: number;
  Title: string;
}

interface IUser {
  Id: number;
  Title: string;
  Email: string;
}

const SharePointGroups: React.FC<IBankManagementProps> = ({ context }) => {

  const navigate = useNavigate();

  const [sp, setSp] = useState<SPFI>();
  const [groups, setGroups] = useState<IGroup[]>([]); 
  const [users, setUsers] = useState<IUser[]>([]);
   const [groupSelected, setGroupSelected] = useState(false);

  useEffect(() => {
    const spInstance = spfi().using(SPFx(context));
    setSp(spInstance);
  }, [context]);

  useEffect(() => {
    if (sp) {
      getGroups();
    }

  }, [sp]);

// all groups getting
  const getGroups = async (): Promise<void> => {

    try {

      const response = await sp!
        .web
        .siteGroups();
      setGroups(response);
      console.log('Groups:', response);

// groups filtered

      // const filteredGroups = response.filter(
      //   (group) => group.Id === 173 || group.Id === 172 || group.Id === 170 || group.Id === 5
      // );
      // setGroups(filteredGroups);

      // ----

      // const allowedGropIds = [173,172,170,5,6];

      // const filteredGroups = response.filter(group =>
      //   allowedGropIds.includes(group.Id)
      // );
      // setGroups(filteredGroups);
      // console.log('Groups:', filteredGroups);
      
      // for (const group of response) {

      //   const users = await sp!
      //   .web
      //   .siteGroups
      //   .getById(group.Id)
      //   .users();

      //   console.log('Group:', group.Title);
      //   console.log('Users:', users);
      // }

    } catch (error) {

      console.log(error);
    }
  };

  //Loads users of clicked group.
  const getUsersByGroup = async (
    groupId: number
  ): Promise<void> => {

    try {

      setGroupSelected(true);

      setUsers([]);

      const response = await sp!
        .web
        .siteGroups
        .getById(groupId)
        .users();

      setUsers(response);
      console.log('Users:', response);


    } catch (error) {

      console.log(error);
    }
  };



  return (

    <div style={{ padding: '20px' }}>


      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '20px'
        }}
      >

        <h1>SharePoint Groups</h1>

        <button
          onClick={() => navigate('/')}
          style={{
            padding: '10px 20px',
            cursor: 'pointer'
          }}
        >
          Back To Home
        </button>

      </div>

      <div
        style={{
          display: 'flex',
          border: '1px solid black',
          minHeight: '500px'
        }}
      >

        {/* LEFT SIDE */}

        <div
          style={{
            width: '50%',
            borderRight: '1px solid black'
          }}
        >

          <h2
            style={{
              textAlign: 'center',
              borderBottom: '1px solid black',
              padding: '10px'
            }}
          >
            Group Names
          </h2>

          {

            groups.map((group) => (

              <div
                key={group.Id}
                onClick={() => getUsersByGroup(group.Id)}
                style={{
                  padding: '12px',
                  cursor: 'pointer',
                  borderBottom: '1px solid #ccc'
                }}
              >
                {group.Title}
              </div>
            ))
          }

        </div>

        {/* RIGHT SIDE */}

        <div
          style={{
            width: '50%'
          }}
        >

          <h2
            style={{
              textAlign: 'center',
              borderBottom: '1px solid black',
              padding: '10px'
            }}
          >
            User Emails
          </h2>

         
          {

            users.length > 0 ? (

              users.map((user) => (

                <div
                  key={user.Id}
                  style={{
                    padding: '12px',
                    borderBottom: '1px solid #ccc'
                  }}
                >
                  {user.Email}
                </div>

              ))

            ) : (

              groupSelected && (

                <div
                  style={{
                    padding: '20px',
                    textAlign: 'center',
                    fontWeight: 'bold'
                  }}
                >
                  Users not exist in this group
                </div>

              )

            )

          }

        </div>

      </div>

    </div>
  );
};

export default SharePointGroups;