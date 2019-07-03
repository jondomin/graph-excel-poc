import React, { useState } from 'react';
import { Table } from 'reactstrap';
import { searchFiles } from '../services/GraphService';
import { Form } from 'reactstrap';
import { FormGroup } from 'reactstrap';
import { Label } from 'reactstrap';
import { Input } from 'reactstrap';
import { Button } from 'reactstrap';
import { Link } from 'react-router-dom';
import LoadingSpinner from '../components/LoadingSpinner';
const OneDriveList = props => {
  const [oneDriveFiles, setOneDriveFiles] = useState([]);
  const [searchText, setSearchText] = useState('');
  const [isLoading, setIsLoading] = useState(false);

  const handleSearch = async evt => {
    evt.preventDefault();
    setIsLoading(true);
    const results = await searchFiles(searchText);
    setIsLoading(false);
    if (results) {
      setOneDriveFiles(results.value);
    }
  };
  return (
    <div>
      <Form controlId="form.InputControl" onSubmit={handleSearch}>
        <FormGroup>
          <Label>Search:</Label>
          <Input
            type="text"
            placeholder="Enter File Names"
            onChange={e => setSearchText(e.target.value)}
          />
        </FormGroup>
        <Button outline color="primary">
          Search
        </Button>
      </Form>
      <Table>
        <LoadingSpinner isLoading={isLoading} />
        <thead>
          <tr>
            <th scope="col">Document Name</th>
          </tr>
        </thead>
        <tbody>
          {oneDriveFiles.map(function(file: any) {
            return (
              <tr key={file.id}>
                <td>
                  <Link to={`${props.match.url}/${file.id}`}>{file.name}</Link>
                </td>
              </tr>
            );
          })}
        </tbody>
      </Table>
    </div>
  );
};

export default OneDriveList;
