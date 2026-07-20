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

    const [sortColumn] =
        React.useState<string>('promptName');

    const [prompts, setPrompts] = React.useState<IPrompt[]>([]);

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
                        gridTemplateColumns: 'repeat(4, 1fr)',
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
                        gap: '10px',
                        flexWrap: 'wrap',
                        marginBottom: '12px'
                    }}
                >
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
                                        <div
                                            style={{
                                                fontWeight: 600,
                                                color: '#0078d4'
                                            }}
                                        >
                                            {item.promptName}
                                        </div>
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
        </div>
    );
};

export default Dashboard;