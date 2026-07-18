import * as React from 'react';

const Header = (): JSX.Element => {
  return (
    <div
      style={{
        background: 'linear-gradient(135deg, #0078d4, #106ebe)',
        color: '#ffffff',
        padding: '24px 32px'
      }}
    >
      <h1
        style={{
          margin: 0,
          fontSize: '32px',
          fontWeight: 700
        }}
      >
        Copilot Prompt Sharing
      </h1>

      <div
        style={{
          marginTop: '8px',
          opacity: 0.9,
          fontSize: '14px'
        }}
      >
        Discover, Share and Reuse AI Prompts Across the Organization
      </div>
    </div>
  );
};

export default Header;