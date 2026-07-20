import * as React from 'react';

export interface IKPICardProps {
  title: string;
  value: number;
  color: string;
  selected?: boolean;
  onClick?: () => void;
}

const KPICard = ({
  title,
  value,
  color,
  selected = false,
  onClick
}: IKPICardProps): JSX.Element => {
  return (
    <div
      onClick={onClick}
      style={{
        backgroundColor: '#ffffff',
        borderRadius: '16px',
        minHeight: '75px',
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'center',
        alignItems: 'center',
        cursor: 'pointer',
        borderLeft: `6px solid ${color}`,
        boxShadow: selected
          ? '0 4px 14px rgba(0,0,0,0.18)'
          : '0 2px 8px rgba(0,0,0,0.08)',
        transform: selected ? 'translateY(-2px)' : 'none',
        transition: 'all 0.2s ease'
      }}
    >
      <div
        style={{
          fontSize: '12px',
          fontWeight: 600,
          color: '#323130',
          marginBottom: '4px'
        }}
      >
        {title}
      </div>

      <div
        style={{
          fontSize: '28px',
          fontWeight: 700,
          color: '#323130'
        }}
      >
        {value}
      </div>
    </div>
  );
};

export default KPICard;