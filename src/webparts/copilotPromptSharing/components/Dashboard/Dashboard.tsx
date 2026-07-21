import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Header from '../Shared/Header';
import KPICard from '../Shared/KPICard';
import PromptService from '../Services/PromptService';
import { IPrompt } from '../Models/IPrompt';

interface IDashboardProps {
    context: WebPartContext;
}

const Dashboard = (props: IDashboardProps): JSX.Element => {

    const [searchText, setSearchText] = React.useState('');
    const [selectedDepartment, setSelectedDepartment] = React.useState('');

    const [isAddPromptOpen, setIsAddPromptOpen] = React.useState(false);

    const [newPrompt, setNewPrompt] = React.useState({
        title: '',
        aiTool: '',
        department: '',
        useCase: '',
        tags: '',
        promptText: ''
    });

    const [sortColumn] =
        React.useState<string>('promptName');

    const [prompts, setPrompts] = React.useState<IPrompt[]>([]);

    const [selectedPrompt, setSelectedPrompt] = React.useState<any>(null);
    const [isModalOpen, setIsModalOpen] = React.useState(false);

    const [sortDirection] =
        React.useState<'asc' | 'desc'>('asc');

    const [selectedStatus, setSelectedStatus] =
        React.useState<string>('All');

    React.useEffect(() => {
        void loadPrompts();
    }, []);

    const loadPrompts = async (): Promise<void> => {
        try {
            const promptService = new PromptService(props.context);
            const items = await promptService.getPrompts();

            setPrompts(items);
        } catch (error) {
            console.log('Error loading prompts:', error);
        } finally {

        }
    };

    const promptData = prompts.map((item: IPrompt) => {
        return {
            id: item.Id,
            promptName: item.Title,
            aiTool: item.AIToolOfChoice,
            department: item.Department,
            status: item.Status,
            tags: item.Tags,
            promptText: item.PromptText,
            useCase: item.UseCase
        };
    });


    const filteredData = promptData.filter((item) => {
        const matchesSearch =
            item.promptName
                .toLowerCase()
                .indexOf(searchText.toLowerCase()) > -1 ||
            item.aiTool
                .toLowerCase()
                .indexOf(searchText.toLowerCase()) > -1;

        const matchesDepartment =
            !selectedDepartment ||
            item.department === selectedDepartment;

        const matchesStatus =
            selectedStatus === 'All' ||
            item.status === selectedStatus;

        return (
            matchesSearch &&
            matchesDepartment &&
            matchesStatus
        );
    });

    const sortedData = [...filteredData];

    sortedData.sort((a, b) => {
        const aValue = String(
            a[sortColumn as keyof typeof a]
        ).toLowerCase();

        const bValue = String(
            b[sortColumn as keyof typeof b]
        ).toLowerCase();

        if (sortDirection === 'asc') {
            return aValue.localeCompare(bValue);
        }

        return bValue.localeCompare(aValue);
    });

    const headerStyle: React.CSSProperties = {
        backgroundColor: '#0078d4',
        color: '#ffffff',
        padding: '14px 16px',
        textAlign: 'left',
        fontWeight: 600,
        fontSize: '14px'
    };

    const cellStyle: React.CSSProperties = {
        padding: '14px 16px',
        borderBottom: '1px solid #edebe9',
        fontSize: '14px'
    };

    const copyPromptText = (): void => {
        if (!selectedPrompt) {
            return;
        }

        navigator.clipboard.writeText(
            selectedPrompt.promptText || ''
        );

        alert('Prompt copied successfully.');
    };

    const savePrompt = async (): Promise<void> => {

        console.log('Prompt To Save', newPrompt);

        // SharePoint save logic will go here next

        setIsAddPromptOpen(false);

        setNewPrompt({
            title: '',
            aiTool: '',
            department: '',
            useCase: '',
            tags: '',
            promptText: ''
        });
    };


    return (
        <div
            style={{
                minHeight: '100vh',
                backgroundColor: '#f3f2f1'
            }}
        >
            <Header />

            <div
                style={{
                    padding: '8px 24px 24px 24px'
                }}
            >
                {/* KPI Cards */}
                <div
                    style={{
                        display: 'grid',
                        gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
                        gap: '16px',
                        marginBottom: '8px'
                    }}
                >
                    <KPICard
                        title="Total Prompts"
                        value={promptData.length}
                        color="#0078D4"
                        selected={selectedStatus === 'All'}
                        onClick={() => setSelectedStatus('All')}
                    />

                    <KPICard
                        title="Approved"
                        value={
                            promptData.filter((item) => item.status === 'Approved').length
                        }
                        color="#107C10"
                        selected={selectedStatus === 'Approved'}
                        onClick={() => setSelectedStatus('Approved')}
                    />

                    <KPICard
                        title="Pending Approval"
                        value={
                            promptData.filter((item) => item.status === 'Send for Approval').length
                        }
                        color="#FF8C00"
                        selected={selectedStatus === 'Send for Approval'}
                        onClick={() => setSelectedStatus('Send for Approval')}
                    />

                    <KPICard
                        title="Rejected"
                        value={
                            promptData.filter((item) => item.status === 'Rejected').length
                        }
                        color="#D13438"
                        selected={selectedStatus === 'Rejected'}
                        onClick={() => setSelectedStatus('Rejected')}
                    />
                </div>

                {/* Search Section */}
                <div
                    style={{
                        background: '#fff',
                        padding: '12px 16px',
                        borderRadius: '12px',
                        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
                        marginBottom: '10px'
                    }}
                >
                    <input
                        type="text"
                        placeholder="Search prompts..."
                        value={searchText}
                        onChange={(e) => setSearchText(e.target.value)}
                        style={{
                            width: '400px',
                            maxWidth: '100%',
                            padding: '12px',
                            border: '1px solid #ddd',
                            borderRadius: '8px',
                            fontSize: '14px'
                        }}
                    />
                </div>

                {/* Department Filters */}
                <div
                    style={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        marginBottom: '12px'
                    }}
                >
                    <button
                        onClick={() => setIsAddPromptOpen(true)}
                        style={{
                            backgroundColor: '#0078d4',
                            color: '#fff',
                            border: 'none',
                            padding: '10px 18px',
                            borderRadius: '8px',
                            cursor: 'pointer',
                            fontWeight: 600,
                            boxShadow: '0 2px 6px rgba(0,0,0,0.15)'
                        }}
                    >
                        + Submit Prompt
                    </button>
                    
                    {['All', 'HR', 'Finance', 'Accounting', 'Stewart AI'].map(
                        (department) => (
                            <button
                                key={department}
                                onClick={() =>
                                    setSelectedDepartment(
                                        department === 'All' ? '' : department
                                    )
                                }

                                style={{
                                    padding: '8px 16px',
                                    borderRadius: '20px',
                                    border: 'none',
                                    cursor: 'pointer',
                                    background:
                                        (
                                            department === 'All' &&
                                            selectedDepartment === ''
                                        ) ||
                                            selectedDepartment === department
                                            ? '#005A9E'
                                            : '#0078D4',
                                    color: '#fff'
                                }}
                            >
                                {department}
                            </button>
                        )
                    )}
                </div>

                <div
                    style={{
                        backgroundColor: '#ffffff',
                        borderRadius: '12px',
                        overflow: 'hidden',
                        boxShadow: '0 2px 8px rgba(0,0,0,0.08)'
                    }}
                >
                    <table
                        style={{
                            width: '100%',
                            borderCollapse: 'collapse'
                        }}
                    >
                        <thead>
                            <tr>
                                <th style={{ ...headerStyle, width: '35%' }}>
                                    Prompt Name
                                </th>

                                <th style={{ ...headerStyle, width: '15%' }}>
                                    Department
                                </th>

                                <th style={{ ...headerStyle, width: '15%' }}>
                                    AI Tool
                                </th>

                                <th style={{ ...headerStyle, width: '15%' }}>
                                    Status
                                </th>

                                <th style={{ ...headerStyle, width: '10%' }}>
                                    Tags
                                </th>

                                <th style={{ ...headerStyle, width: '10%' }}>
                                    Actions
                                </th>
                            </tr>
                        </thead>

                        <tbody>
                            {sortedData.map((item, index) => (
                                <tr
                                    key={item.id}
                                    style={{
                                        backgroundColor:
                                            index % 2 === 0
                                                ? '#ffffff'
                                                : '#f8f9fb'
                                    }}
                                >
                                    <td style={cellStyle}>
                                        <a
                                            href="#"
                                            onClick={(e) => {
                                                e.preventDefault();

                                                console.log('Clicked Prompt:', item);

                                                setSelectedPrompt(item);
                                                setIsModalOpen(true);
                                            }}
                                            style={{
                                                color: '#0078d4',
                                                fontWeight: 600,
                                                textDecoration: 'underline',
                                                cursor: 'pointer'
                                            }}
                                        >
                                            {item.promptName}
                                        </a>
                                    </td>

                                    <td style={cellStyle}>
                                        {item.department}
                                    </td>

                                    <td style={cellStyle}>
                                        {item.aiTool}
                                    </td>

                                    <td style={cellStyle}>
                                        <span
                                            style={{
                                                padding: '4px 10px',
                                                borderRadius: '20px',
                                                fontWeight: 600,
                                                whiteSpace: 'nowrap',
                                                background:
                                                    item.status === 'Approved'
                                                        ? '#dff6dd'
                                                        : item.status === 'Rejected'
                                                            ? '#fde7e9'
                                                            : '#fff4ce',
                                                color:
                                                    item.status === 'Approved'
                                                        ? '#107c10'
                                                        : item.status === 'Rejected'
                                                            ? '#d13438'
                                                            : '#8a6d3b'
                                            }}
                                        >
                                            {item.status}
                                        </span>
                                    </td>

                                    <td style={cellStyle}>
                                        {item.tags}
                                    </td>

                                    <td style={cellStyle}>
                                        <span
                                            style={{
                                                color: '#0078d4',
                                                cursor: 'pointer',
                                                fontWeight: 600
                                            }}
                                        >
                                            View
                                        </span>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
            {isModalOpen && selectedPrompt && (
                <div
                    style={{
                        position: 'fixed',
                        top: 0,
                        left: 0,
                        right: 0,
                        bottom: 0,
                        backgroundColor: 'rgba(0,0,0,0.5)',
                        zIndex: 9999,
                        overflowY: 'auto',
                        padding: '30px'
                    }}
                >
                    <div
                        style={{
                            width: '800px',
                            maxWidth: '95%',
                            margin: '0 auto',
                            backgroundColor: '#ffffff',
                            borderRadius: '18px',
                            overflow: 'hidden',
                            boxShadow: '0 10px 30px rgba(0,0,0,0.25)'
                        }}
                    >
                        {/* HEADER */}
                        <div
                            style={{
                                background: 'linear-gradient(135deg,#006d8f,#1284b3)',
                                color: '#fff',
                                padding: '28px 36px',
                                position: 'relative'
                            }}
                        >
                            <button
                                onClick={() => setIsModalOpen(false)}
                                style={{
                                    position: 'absolute',
                                    top: '20px',
                                    right: '20px',
                                    width: '42px',
                                    height: '42px',
                                    borderRadius: '50%',
                                    border: '1px solid rgba(255,255,255,.3)',
                                    background: 'transparent',
                                    color: '#fff',
                                    cursor: 'pointer',
                                    fontSize: '22px'
                                }}
                            >
                                ✕
                            </button>

                            <h1
                                style={{
                                    margin: 0,
                                    fontSize: '40px',
                                    fontWeight: 700,
                                    lineHeight: '1.2'
                                }}
                            >
                                {selectedPrompt.promptName}
                            </h1>
                        </div>

                        {/* BODY */}
                        <div
                            style={{
                                padding: '24px',
                                backgroundColor: '#f5f7fa'
                            }}
                        >
                            {/* USE CASE */}
                            <div
                                style={{
                                    background: '#fff',
                                    borderRadius: '12px',
                                    padding: '20px',
                                    marginBottom: '20px'
                                }}
                            >
                                <div
                                    style={{
                                        color: '#006d8f',
                                        fontWeight: 700,
                                        marginBottom: '12px',
                                        letterSpacing: '1px'
                                    }}
                                >
                                    USE CASE
                                </div>

                                {selectedPrompt.useCase}
                            </div>

                            {/* DETAILS */}
                            <div
                                style={{
                                    background: '#fff',
                                    borderRadius: '12px',
                                    padding: '20px',
                                    marginBottom: '20px'
                                }}
                            >
                                <div
                                    style={{
                                        color: '#006d8f',
                                        fontWeight: 700,
                                        marginBottom: '20px',
                                        letterSpacing: '1px'
                                    }}
                                >
                                    DETAILS
                                </div>

                                <div
                                    style={{
                                        display: 'grid',
                                        gridTemplateColumns: '1fr 1fr',
                                        gap: '24px'
                                    }}
                                >
                                    <div>
                                        <strong>Department</strong>
                                        <div>{selectedPrompt.department}</div>
                                    </div>

                                    <div>
                                        <strong>AI Tool</strong>
                                        <div>{selectedPrompt.aiTool}</div>
                                    </div>

                                    <div>
                                        <strong>Status</strong>
                                        <div>{selectedPrompt.status}</div>
                                    </div>

                                    <div>
                                        <strong>Tags</strong>
                                        <div>{selectedPrompt.tags}</div>
                                    </div>
                                </div>
                            </div>

                            {/* PROMPT TEXT */}
                            <div
                                style={{
                                    background: '#fff',
                                    borderRadius: '12px',
                                    padding: '20px'
                                }}
                            >
                                <div
                                    style={{
                                        display: 'flex',
                                        justifyContent: 'space-between',
                                        alignItems: 'center',
                                        marginBottom: '15px'
                                    }}
                                >
                                    <div
                                        style={{
                                            color: '#006d8f',
                                            fontWeight: 700,
                                            letterSpacing: '1px'
                                        }}
                                    >
                                        PROMPT TEXT
                                    </div>

                                    <span
                                        onClick={copyPromptText}
                                        title="Copy Prompt"
                                        style={{
                                            cursor: 'pointer',
                                            fontSize: '22px'
                                        }}
                                    >
                                        📋
                                    </span>
                                </div>

                                <div
                                    style={{
                                        background: '#f8f9fb',
                                        padding: '18px',
                                        borderRadius: '8px',
                                        border: '1px solid #edebe9',
                                        whiteSpace: 'pre-wrap',
                                        lineHeight: '1.7'
                                    }}
                                >
                                    {selectedPrompt.promptText}
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            )}
            {isAddPromptOpen && (
                <div
                    style={{
                        position: 'fixed',
                        inset: 0,
                        backgroundColor: 'rgba(0,0,0,.5)',
                        zIndex: 9999,
                        overflowY: 'auto',
                        padding: '30px'
                    }}
                >
                    <div
                        style={{
                            width: '900px',
                            maxWidth: '95%',
                            margin: '0 auto',
                            backgroundColor: '#fff',
                            borderRadius: '18px',
                            overflow: 'hidden'
                        }}
                    >
                        <div
                            style={{
                                background: 'linear-gradient(135deg,#006d8f,#1284b3)',
                                color: '#fff',
                                padding: '24px 32px',
                                position: 'relative'
                            }}
                        >
                            <button
                                onClick={() => setIsAddPromptOpen(false)}
                                style={{
                                    position: 'absolute',
                                    top: '20px',
                                    right: '20px',
                                    width: '42px',
                                    height: '42px',
                                    borderRadius: '50%',
                                    border: '1px solid rgba(255,255,255,.3)',
                                    background: 'transparent',
                                    color: '#fff',
                                    cursor: 'pointer',
                                    fontSize: '22px'
                                }}
                            >
                                ✕
                            </button>

                            <h2
                                style={{
                                    margin: 0
                                }}
                            >
                                Submit New Prompt
                            </h2>
                        </div>

                        <div
                            style={{
                                padding: '24px'
                            }}
                        >

                            <div style={{ marginBottom: '16px' }}>
                                <label>Prompt Name</label>
                                <input
                                    value={newPrompt.title}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            title: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div style={{ marginBottom: '16px' }}>
                                <label>AI Tool</label>
                                <input
                                    value={newPrompt.aiTool}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            aiTool: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div style={{ marginBottom: '16px' }}>
                                <label>Department</label>
                                <input
                                    value={newPrompt.department}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            department: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div style={{ marginBottom: '16px' }}>
                                <label>Use Case</label>
                                <textarea
                                    rows={3}
                                    value={newPrompt.useCase}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            useCase: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div style={{ marginBottom: '16px' }}>
                                <label>Tags</label>
                                <input
                                    value={newPrompt.tags}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            tags: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div style={{ marginBottom: '16px' }}>
                                <label>Prompt Text</label>
                                <textarea
                                    rows={8}
                                    value={newPrompt.promptText}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            promptText: e.target.value
                                        })
                                    }
                                    style={{
                                        width: '100%',
                                        padding: '10px'
                                    }}
                                />
                            </div>

                            <div
                                style={{
                                    display: 'flex',
                                    justifyContent: 'flex-end',
                                    gap: '10px'
                                }}
                            >
                                <button
                                    onClick={() => setIsAddPromptOpen(false)}
                                >
                                    Cancel
                                </button>

                                <button
                                    onClick={savePrompt}
                                    style={{
                                        backgroundColor: '#0078d4',
                                        color: '#fff',
                                        border: 'none',
                                        padding: '10px 18px',
                                        borderRadius: '6px'
                                    }}
                                >
                                    Submit
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </div >
    );
};

export default Dashboard;