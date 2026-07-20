import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Header from '../Shared/Header';
import KPICard from '../Shared/KPICard';
import PromptService from '../Services/PromptService';
import { IPrompt } from '../Models/IPrompt';
import {
    DetailsList,
    DetailsListLayoutMode,
    SelectionMode,
    IColumn
} from '@fluentui/react';

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


    const paginationButtonStyle: React.CSSProperties = {
        padding: '6px 12px',
        marginLeft: '8px',
        borderRadius: '6px',
        border: '1px solid #d1d1d1',
        background: '#ffffff',
        cursor: 'pointer'
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


    const columns: IColumn[] = [
        {
            key: 'promptName',
            name: 'Prompt Name',
            fieldName: 'promptName',
            minWidth: 250,
            maxWidth: 350,
            isResizable: true
        },
        {
            key: 'department',
            name: 'Department',
            fieldName: 'department',
            minWidth: 120,
            maxWidth: 150
        },
        {
            key: 'aiTool',
            name: 'AI Tool',
            fieldName: 'aiTool',
            minWidth: 120,
            maxWidth: 150
        },
        {
            key: 'status',
            name: 'Status',
            fieldName: 'status',
            minWidth: 140,
            maxWidth: 180,
            onRender: (item) => {
                const background =
                    item.status === 'Approved'
                        ? '#dff6dd'
                        : item.status === 'Rejected'
                            ? '#fde7e9'
                            : '#fff4ce';

                const color =
                    item.status === 'Approved'
                        ? '#107c10'
                        : item.status === 'Rejected'
                            ? '#d13438'
                            : '#8a6d3b';

                return (
                    <span
                        style={{
                            padding: '4px 10px',
                            borderRadius: '12px',
                            background,
                            color,
                            fontWeight: 600
                        }}
                    >
                        {item.status}
                    </span>
                );
            }
        },
        {
            key: 'tags',
            name: 'Tags',
            fieldName: 'tags',
            minWidth: 150
        },
        {
            key: 'actions',
            name: 'Actions',
            minWidth: 120,
            onRender: () => (
                <>
                    <button
                        style={{
                            border: 'none',
                            background: 'transparent',
                            color: '#0078d4',
                            cursor: 'pointer',
                            marginRight: '12px'
                        }}
                    >
                        View
                    </button>

                    <button
                        style={{
                            border: 'none',
                            background: 'transparent',
                            color: '#0078d4',
                            cursor: 'pointer'
                        }}
                    >
                        Copy
                    </button>
                </>
            )
        }
    ];


    return (
        <div
            style={{
                minHeight: '100vh',
                backgroundColor: '#f5f7fa'
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

                <DetailsList
                    items={sortedData}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    layoutMode={DetailsListLayoutMode.justified}
                    compact={true}
                />


                <span>Showing {sortedData.length} prompt(s)</span>

                <div>
                    <button style={paginationButtonStyle}>Previous</button>
                    <button style={paginationButtonStyle}>Next</button>
                </div>
            </div>
        </div>
    );
};

export default Dashboard;