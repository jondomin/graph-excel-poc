import React from 'react';
import { Spinner } from 'reactstrap';

const LoadingSpinner = ({ isLoading }) => {
  return isLoading && <Spinner size="md" color="primary" />;
};

export default LoadingSpinner;
