import React, { useState, useEffect } from 'react';
import { Table } from 'reactstrap';
import { getWorksheets } from '../services/GraphService';
import { Link } from 'react-router-dom';
const OneDrive = props => {
  const [worksheets, setWorksheets] = useState([]);
  const msal = window.msal as any;

  useEffect(() => {
    async function loadSheets() {
      try {
        const result = await getWorksheets(props.match.params.documentId);
        setWorksheets(result.value);
      } catch (error) {
        props.showError('ERROR', JSON.stringify(error));
      }
    }

    loadSheets();
    return () => {};
  }, [msal, props]);
  return (
    <div>
      <Table>
        <thead>
          <tr>
            <th scope="col">Sheet Name</th>
          </tr>
        </thead>
        <tbody>
          {worksheets.map(function(worksheet: any) {
            return (
              <tr key={worksheet.id}>
                <td>
                  <Link to={`/table/${props.match.params.documentId}/${worksheet.name}/`}>
                    {worksheet.name}
                  </Link>
                </td>
              </tr>
            );
          })}
        </tbody>
      </Table>
    </div>
  );
};

export default OneDrive;
