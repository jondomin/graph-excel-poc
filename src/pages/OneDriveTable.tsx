import React, { useState, useEffect } from 'react';
import { Table, Row, Col } from 'reactstrap';
import { getTable } from '../services/GraphService';
import './OneDriveTable.scss';
import LoadingSpinner from '../components/LoadingSpinner';
const OneDriveTable = props => {
  const [tableData, setTableData] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const msal = window.msal as any;

  useEffect(() => {
    async function loadTable() {
      try {
        setIsLoading(true);
        const result = await getTable(props.match.params.documentId, props.match.params.sheetName);
        setIsLoading(false);
        setTableData(result.values);
      } catch (error) {
        props.showError('ERROR', JSON.stringify(error));
      }
    }

    loadTable();
    return () => {};
  }, [msal, props]);
  return (
    <Row>
      <Col>
        <h1>{props.match.params.sheetName}</h1>
        <Row>
          <Col className="onedrivetable__wrapper">
            <LoadingSpinner isLoading={isLoading} />
            <Table striped>
              {tableData.map(function(row: [], index: number) {
                if (index === 0) {
                  return (
                    <thead>
                      <tr>
                        {row.map(function(cell) {
                          return <th>{cell}</th>;
                        })}
                      </tr>
                    </thead>
                  );
                } else {
                  return (
                    <tbody>
                      <tr>
                        {row.map(function(cell) {
                          return <td>{cell}</td>;
                        })}
                      </tr>
                    </tbody>
                  );
                }
              })}
            </Table>
          </Col>
        </Row>
      </Col>
    </Row>
  );
};

export default OneDriveTable;
