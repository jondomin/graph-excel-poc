import React, { useState, useEffect } from 'react';
import { Table, Row, Col } from 'reactstrap';
import { getWorksheets } from '../services/GraphService';
import { Link } from 'react-router-dom';
import LoadingSpinner from '../components/LoadingSpinner';
const OneDrive = props => {
  const [worksheets, setWorksheets] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const msal = window.msal as any;

  useEffect(() => {
    async function loadSheets() {
      try {
        setIsLoading(true);
        const result = await getWorksheets(props.match.params.documentId);
        setIsLoading(false);
        setWorksheets(result.value);
      } catch (error) {
        props.showError('ERROR', JSON.stringify(error));
      }
    }

    loadSheets();
    return () => {};
  }, [msal, props]);
  return (
    <Row>
      <Col>
        <LoadingSpinner isLoading={isLoading} />
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
      </Col>
    </Row>
  );
};

export default OneDrive;
