import * as React from 'react';

const Header = (): JSX.Element => {
    return (
        <div
            style={{
                background: 'linear-gradient(135deg, #066C90,#0A83AE)',
                padding: '18px 24px',
                color: '#fff',
                borderRadius: '18px',
                margin: '16px',
                textAlign: 'center',
                boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
                marginBottom: '10px'
            }}
        >
            <h1
                style={{
                    margin: 0,
                    fontSize: '24px',
                    fontWeight: 700
                }}
            >
                Stewart AI Prompt Library
            </h1>

            <div
                style={{
                    marginTop: '4px',
                    opacity: 0.9
                }}
            >
                Discover, Share and Reuse AI Prompts Across Stewart Title
            </div>
        </div>
    );
};

export default Header;