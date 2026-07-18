import * as React from 'react';
import Header from '../Shared/Header';

const Dashboard = (): JSX.Element => {
  return (
    <div
      style={{
        minHeight: '100vh',
        backgroundColor: '#f5f7fa'
      }}
    >
      <Header />

      <div style={{ padding: '24px' }}>
        Dashboard Content Coming Soon...
      </div>
    </div>
  );
};

export default Dashboard;