import * as React from 'react';

export interface IKPICardProps {
  title: string;
  value: string | number;
}

const KPICard = ({
  title,
  value
}: IKPICardProps): JSX.Element => {
  return (
    <div
      style={{
        background: '#ffffff',
        borderRadius: '12px',
        padding: '20px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
        minWidth: '220px'
      }}
    >
      <div
        style={{
          color: '#605e5c',
          fontSize: '14px'
        }}
      >
        {title}
      </div>

      <div
        style={{
          marginTop: '8px',
          fontSize: '28px',
          fontWeight: 700,
          color: '#0078d4'
        }}
      >
        {value}
      </div>
    </div>
  );
};

export default KPICard;