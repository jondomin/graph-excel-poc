import React, { useState, useEffect } from 'react';
import { Table, Row, Col } from 'reactstrap';
import { getTable } from '../services/GraphService';
import './OneDriveTable.scss';
const OneDriveTable = props => {
  const [tableData, setTableData] = useState([]);
  const msal = window.msal as any;

  useEffect(() => {
    async function loadTable() {
      try {
        const result = await getTable(props.match.params.documentId, props.match.params.sheetName);
        debugger;
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
            <Table>
              <tbody>
                {tableData.map(function(row: []) {
                  return (
                    <tr>
                      {row.map(function(cell) {
                        return <td>{cell}</td>;
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </Table>
          </Col>
        </Row>
      </Col>
    </Row>
  );
};

export default OneDriveTable;
