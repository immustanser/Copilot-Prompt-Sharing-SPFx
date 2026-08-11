import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import Header from '../Shared/Header';
import KPICard from '../Shared/KPICard';
import PromptService from '../Services/PromptService';
import { IPrompt } from '../Models/IPrompt';

interface IDashboardProps {
    context: WebPartContext;
}

interface IToast {
    id: number;
    message: string;
    type: 'success' | 'error';
}

const Dashboard = (props: IDashboardProps): JSX.Element => {

    // Single service instance for the component lifetime
    const promptService = React.useMemo(
        () => new PromptService(props.context),
        []
    );

    // Cached after first load – reused by post-CRUD lightweight refreshes
    const currentUserIdRef = React.useRef<number>(0);

    const toastIdRef = React.useRef(0);
    const [toasts, setToasts] = React.useState<IToast[]>([]);

    const showToast = (
        message: string,
        type: 'success' | 'error' = 'success'
    ): void => {
        const id = ++toastIdRef.current;
        setToasts(prev => [...prev, { id, message, type }]);
        setTimeout(() => {
            setToasts(prev => prev.filter(t => t.id !== id));
        }, 4500);
    };

    const dismissToast = (id: number): void => {
        setToasts(prev => prev.filter(t => t.id !== id));
    };

    const [searchText, setSearchText] = React.useState('');
    const [selectedDepartment, setSelectedDepartment] = React.useState('');

    const [isAddPromptOpen, setIsAddPromptOpen] = React.useState(false);
    const [isEditMode, setIsEditMode] = React.useState(false);
    const [editPromptId, setEditPromptId] = React.useState<number | null>(null);

    const [currentPage, setCurrentPage] = React.useState(1);
    const itemsPerPage = 10;


    const [newPrompt, setNewPrompt] = React.useState({
        title: '',
        aiTool: '',
        department: '',
        useCase: '',
        tags: '',
        promptText: ''
    });

    const [sortColumn, setSortColumn] =
        React.useState<string>('promptName');

    const [sortDirection, setSortDirection] =
        React.useState<'asc' | 'desc'>('asc');

    const [prompts, setPrompts] = React.useState<IPrompt[]>([]);

    const [selectedPrompt, setSelectedPrompt] = React.useState<any>(null);
    const [isModalOpen, setIsModalOpen] = React.useState(false);

    const [canAdd, setCanAdd] = React.useState(false);
    const [canEdit, setCanEdit] = React.useState(false);
    const [canDelete, setCanDelete] = React.useState(false);

    const [departmentOptions, setDepartmentOptions] =
        React.useState<string[]>([]);

    const [aiToolOptions, setAiToolOptions] =
        React.useState<string[]>([]);

    const [isPageLoading, setIsPageLoading] = React.useState(true);
    const [isSaving, setIsSaving] = React.useState(false);
    const [deletingId, setDeletingId] = React.useState<number | null>(null);

    const [selectedStatus, setSelectedStatus] =
        React.useState<string>('All');

    React.useEffect(() => {
        void loadPrompts();
    }, []);

    React.useEffect(() => {
        setCurrentPage(1);
    }, [
        searchText,
        selectedDepartment,
        selectedStatus
    ]);


    const loadPrompts = async (): Promise<void> => {
        try {
            const data = await promptService.initializeDashboard();
            currentUserIdRef.current = data.currentUserId;
            setPrompts(data.prompts);
            setDepartmentOptions(data.departmentChoices);
            setAiToolOptions(data.aiToolChoices);
            setCanAdd(data.permissions.canAdd);
            setCanEdit(data.permissions.canEdit);
            setCanDelete(data.permissions.canDelete);
        } catch (error) {
            console.error('Error loading dashboard:', error);
            showToast('Failed to load prompts. Please refresh the page.', 'error');
        } finally {
            setIsPageLoading(false);
        }
    };

    const promptData = prompts.map((item: IPrompt) => {
        return {
            id: item.Id,
            promptName: item.Title,
            aiTool: item.AiTool,
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
            a[sortColumn as keyof typeof a] || ''
        ).toLowerCase();

        const bValue = String(
            b[sortColumn as keyof typeof b] || ''
        ).toLowerCase();

        if (sortDirection === 'asc') {
            return aValue.localeCompare(bValue);
        }

        return bValue.localeCompare(aValue);
    });

    const totalPages = Math.ceil(
        sortedData.length / itemsPerPage
    );

    const paginatedData = sortedData.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage
    );

    const headerStyle: React.CSSProperties = {
        backgroundColor: '#0A83AE',
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


    const truncateStyle: React.CSSProperties = {
        maxWidth: '180px',
        overflow: 'hidden',
        textOverflow: 'ellipsis',
        whiteSpace: 'nowrap'
    };
    const copyPromptText = (): void => {
        if (!selectedPrompt) {
            return;
        }

        navigator.clipboard.writeText(
            selectedPrompt.promptText || ''
        );

        showToast('Prompt copied to clipboard.');
    };

    const openAddPromptModal = (): void => {
        setIsEditMode(false);
        setEditPromptId(null);

        setNewPrompt({
            title: '',
            aiTool: '',
            department: '',
            useCase: '',
            tags: '',
            promptText: ''
        });

        setIsAddPromptOpen(true);
    };

    const openEditPromptModal = (item: any): void => {
        setIsEditMode(true);
        setEditPromptId(item.id);

        setNewPrompt({
            title: item.promptName || '',
            aiTool: item.aiTool || '',
            department: item.department || '',
            useCase: item.useCase || '',
            tags: item.tags || '',
            promptText: item.promptText || ''
        });

        setIsAddPromptOpen(true);
    };

    const closeAddPromptModal = (): void => {
        setIsAddPromptOpen(false);
        setIsEditMode(false);
        setEditPromptId(null);

        setNewPrompt({
            title: '',
            aiTool: '',
            department: '',
            useCase: '',
            tags: '',
            promptText: ''
        });
    };

    const savePrompt = async (): Promise<void> => {
        const title = newPrompt.title.trim();
        const aiTool = newPrompt.aiTool.trim();
        const department = newPrompt.department.trim();
        const useCase = newPrompt.useCase.trim();
        const tags = newPrompt.tags.trim();
        const promptText = newPrompt.promptText.trim();

        if (!title) {
            showToast('Please enter Prompt Name.', 'error');
            return;
        }

        if (!aiTool) {
            showToast('Please select AI Tool.', 'error');
            return;
        }

        if (!department) {
            showToast('Please select Department.', 'error');
            return;
        }

        if (!useCase) {
            showToast('Please enter Use Case.', 'error');
            return;
        }

        if (!promptText) {
            showToast('Please enter Prompt Text.', 'error');
            return;
        }

        setIsSaving(true);

        try {
            if (isEditMode && editPromptId !== null) {
                await promptService.updatePrompt(editPromptId, {
                    title, aiTool, department, useCase, tags, promptText
                });

                // Optimistic update — all values are known, no network refresh needed
                setPrompts(prev => prev.map(p =>
                    p.Id === editPromptId
                        ? {
                            Id: editPromptId,
                            Title: title,
                            AiTool: aiTool,
                            Department: department,
                            UseCase: useCase,
                            Tags: tags,
                            PromptText: promptText,
                            Status: 'Send for Approval'
                        }
                        : p
                ));

                closeAddPromptModal();
                setSelectedStatus('Send for Approval');
                setSelectedDepartment('');
                showToast('Prompt updated successfully and sent for approval.');

            } else {
                await promptService.createPrompt({
                    title, aiTool, department, useCase, tags, promptText
                });

                // Close modal and notify immediately for snappy UX
                closeAddPromptModal();
                setSelectedStatus('Send for Approval');
                setSelectedDepartment('');
                showToast('Prompt submitted successfully and sent for approval.');

                // Lightweight refresh to get the new item with its server-assigned Id
                const freshItems = await promptService.refreshPrompts(
                    currentUserIdRef.current
                );
                setPrompts(freshItems);
            }
        } catch (error) {
            console.error('Error saving prompt:', error);
            showToast(
                'Something went wrong while saving the prompt. Please try again.',
                'error'
            );
        } finally {
            setIsSaving(false);
        }
    };

    const departments = ['All'];

    promptData.forEach(item => {
        if (
            item.department &&
            departments.indexOf(item.department) === -1
        ) {
            departments.push(item.department);
        }
    });


    const deletePrompt = async (itemId: number): Promise<void> => {
        const confirmed = window.confirm(
            'Are you sure you want to delete this prompt?'
        );
        if (!confirmed) return;

        setDeletingId(itemId);

        try {
            await promptService.deletePrompt(itemId);

            // Optimistic remove — no round trip needed since we know exactly what was deleted
            setPrompts(prev => prev.filter(p => p.Id !== itemId));
            showToast('Prompt deleted successfully.');

        } catch (error) {
            console.error('Error deleting prompt:', error);
            showToast('Unable to delete prompt. Please try again.', 'error');
        } finally {
            setDeletingId(null);
        }
    };

    const handleSort = (
        column: string
    ): void => {

        if (sortColumn === column) {

            setSortDirection(
                sortDirection === 'asc'
                    ? 'desc'
                    : 'asc'
            );

        } else {

            setSortColumn(column);
            setSortDirection('asc');

        }
    };

    return (
        <div
            style={{
                minHeight: '100vh',
                backgroundColor: '#f3f2f1'
            }}
        >
            <style>{`
                @keyframes cpsp-spin {
                    to { transform: rotate(360deg); }
                }
                @keyframes cpsp-fadeIn {
                    from { opacity: 0; transform: translateX(24px); }
                    to   { opacity: 1; transform: translateX(0); }
                }
            `}</style>

            {isPageLoading && (
                <div
                    style={{
                        position: 'fixed',
                        inset: 0,
                        backgroundColor: 'rgba(243,242,241,0.97)',
                        zIndex: 9998,
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '18px'
                    }}
                >
                    <div
                        style={{
                            width: '52px',
                            height: '52px',
                            border: '5px solid rgba(10,131,174,0.2)',
                            borderTopColor: '#0A83AE',
                            borderRadius: '50%',
                            animation: 'cpsp-spin 0.75s linear infinite'
                        }}
                    />
                    <p
                        style={{
                            margin: 0,
                            color: '#0A83AE',
                            fontWeight: 600,
                            fontSize: '16px',
                            letterSpacing: '0.3px'
                        }}
                    >
                        Loading prompts...
                    </p>
                </div>
            )}

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
                        marginBottom: '10px'
                    }}
                >
                    <KPICard
                        title="Total Prompts"
                        value={promptData.length}
                        color="#0A83AE"
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
                        padding: '16px',
                        borderRadius: '12px',
                        boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
                        marginBottom: '12px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '16px'
                    }}
                >
                    <button
                        disabled={!canAdd}
                        onClick={
                            canAdd
                                ? openAddPromptModal
                                : undefined
                        }
                        title={
                            !canAdd
                                ? 'You do not have permissions to add prompts'
                                : ''
                        }
                        style={{
                            backgroundColor: '#0A83AE',
                            color: '#fff',
                            border: 'none',
                            padding: '10px 18px',
                            borderRadius: '8px',
                            cursor: canAdd
                                ? 'pointer'
                                : 'not-allowed',
                            fontWeight: 600,
                            whiteSpace: 'nowrap'
                        }}
                    >
                        + Add New Prompt
                    </button>

                    <input
                        type="text"
                        placeholder="Search prompts..."
                        value={searchText}
                        onChange={(e) =>
                            setSearchText(e.target.value)
                        }
                        style={{
                            flex: 1,
                            padding: '12px',
                            border: '1px solid #ddd',
                            borderRadius: '8px',
                            fontSize: '14px'
                        }}
                    />

                    <div
                        style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            whiteSpace: 'nowrap'
                        }}
                    >
                        <span
                            style={{
                                fontWeight: 600,
                                color: '#323130'
                            }}
                        >
                            Department:
                        </span>

                        <select
                            value={selectedDepartment}
                            onChange={(e) =>
                                setSelectedDepartment(
                                    e.target.value
                                )
                            }
                            style={{
                                padding: '10px',
                                borderRadius: '8px',
                                border: '1px solid #d1d1d1',
                                minWidth: '180px'
                            }}
                        >
                            <option value="">
                                All
                            </option>

                            {departments
                                .filter(
                                    department =>
                                        department !== 'All'
                                )
                                .map(
                                    (department) => (
                                        <option
                                            key={department}
                                            value={department}
                                        >
                                            {department}
                                        </option>
                                    )
                                )}
                        </select>
                    </div>
                </div>

                <div
                    style={{
                        backgroundColor: '#ffffff',
                        borderRadius: '12px',
                        overflowX: 'auto',
                        overflowY: 'hidden',
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
                                <th
                                    onClick={() =>
                                        handleSort('promptName')
                                    }
                                    style={{
                                        ...headerStyle,
                                        width: '35%',
                                        cursor: 'pointer'
                                    }}
                                >
                                    Prompt Name
                                    {
                                        sortColumn === 'promptName'
                                            ? sortDirection === 'asc'
                                                ? ' ▲'
                                                : ' ▼'
                                            : ''
                                    }
                                </th>

                                <th
                                    onClick={() =>
                                        handleSort('department')
                                    }
                                    style={{
                                        ...headerStyle,
                                        width: '15%',
                                        cursor: 'pointer'
                                    }}
                                >
                                    Department
                                    {
                                        sortColumn === 'department'
                                            ? sortDirection === 'asc'
                                                ? ' ▲'
                                                : ' ▼'
                                            : ''
                                    }
                                </th>

                                <th
                                    onClick={() =>
                                        handleSort('aiTool')
                                    }
                                    style={{
                                        ...headerStyle,
                                        width: '15%',
                                        cursor: 'pointer'
                                    }}
                                >
                                    AI Tool
                                    {
                                        sortColumn === 'aiTool'
                                            ? sortDirection === 'asc'
                                                ? ' ▲'
                                                : ' ▼'
                                            : ''
                                    }
                                </th>

                                <th
                                    onClick={() =>
                                        handleSort('status')
                                    }
                                    style={{
                                        ...headerStyle,
                                        width: '15%',
                                        cursor: 'pointer'
                                    }}
                                >
                                    Status
                                    {
                                        sortColumn === 'status'
                                            ? sortDirection === 'asc'
                                                ? ' ▲'
                                                : ' ▼'
                                            : ''
                                    }
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
                            {paginatedData.map((item, index) => (
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
                                                color: '#0A83AE',
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
                                        <div
                                            title={item.aiTool}
                                            style={{
                                                ...truncateStyle,
                                                maxWidth: '90px'
                                            }}
                                        >
                                            {item.aiTool}
                                        </div>
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
                                        <div
                                            title={item.tags}
                                            style={{
                                                ...truncateStyle,
                                                maxWidth: '90px'
                                            }}
                                        >
                                            {item.tags}
                                        </div>
                                    </td>

                                    <td style={cellStyle}>
                                        <div
                                            style={{
                                                display: 'flex',
                                                gap: '12px',
                                                alignItems: 'center'
                                            }}
                                        >
                                            {item.status !== 'Send for Approval' && (
                                                deletingId === item.id ? (
                                                    <span
                                                        style={{
                                                            display: 'flex',
                                                            alignItems: 'center',
                                                            gap: '5px',
                                                            fontSize: '12px',
                                                            color: '#0A83AE',
                                                            fontWeight: 600
                                                        }}
                                                    >
                                                        <span
                                                            style={{
                                                                width: '14px',
                                                                height: '14px',
                                                                border: '2px solid rgba(10,131,174,0.25)',
                                                                borderTopColor: '#0A83AE',
                                                                borderRadius: '50%',
                                                                display: 'inline-block',
                                                                flexShrink: 0,
                                                                animation: 'cpsp-spin 0.75s linear infinite'
                                                            }}
                                                        />
                                                        Deleting...
                                                    </span>
                                                ) : (
                                                    <>
                                                        <span
                                                            onClick={
                                                                canEdit && deletingId === null
                                                                    ? () => openEditPromptModal(item)
                                                                    : undefined
                                                            }
                                                            style={{
                                                                opacity: canEdit && deletingId === null ? 1 : 0.4,
                                                                cursor: canEdit && deletingId === null ? 'pointer' : 'not-allowed'
                                                            }}
                                                            title={
                                                                !canEdit
                                                                    ? 'You do not have permissions to edit prompts'
                                                                    : ''
                                                            }
                                                        >
                                                            ✏️
                                                        </span>

                                                        <span
                                                            onClick={
                                                                canDelete && deletingId === null
                                                                    ? () => deletePrompt(item.id)
                                                                    : undefined
                                                            }
                                                            style={{
                                                                opacity: canDelete && deletingId === null ? 1 : 0.4,
                                                                cursor: canDelete && deletingId === null ? 'pointer' : 'not-allowed'
                                                            }}
                                                            title={
                                                                !canDelete
                                                                    ? 'You do not have permissions to delete prompts'
                                                                    : ''
                                                            }
                                                        >
                                                            🗑️
                                                        </span>
                                                    </>
                                                )
                                            )}
                                        </div>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                    <div
                        style={{
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center',
                            padding: '12px 20px',
                            borderTop: '1px solid #edebe9',
                            backgroundColor: '#fafafa'
                        }}
                    >
                        <span>
                            Showing {
                                paginatedData.length > 0
                                    ? ((currentPage - 1) * itemsPerPage) + 1
                                    : 0
                            }
                            -
                            {
                                ((currentPage - 1) * itemsPerPage) +
                                paginatedData.length
                            }
                            {' '}of{' '}
                            {sortedData.length}
                        </span>

                        <div
                            style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: '6px'
                            }}
                        >
                            <span
                                onClick={() => setCurrentPage(1)}
                                style={{
                                    cursor: currentPage === 1 ? 'not-allowed' : 'pointer',
                                    opacity: currentPage === 1 ? 0.4 : 1,
                                    padding: '6px',
                                    userSelect: 'none'
                                }}
                            >
                                «
                            </span>

                            <span
                                onClick={() => {
                                    if (currentPage > 1) {
                                        setCurrentPage(currentPage - 1);
                                    }
                                }}
                                style={{
                                    cursor: currentPage === 1 ? 'not-allowed' : 'pointer',
                                    opacity: currentPage === 1 ? 0.4 : 1,
                                    padding: '6px',
                                    userSelect: 'none'
                                }}
                            >
                                ‹
                            </span>

                            {[...Array(totalPages)].map(
                                (_item: undefined, index: number) => {
                                    const pageNumber = index + 1;

                                    return (
                                        <span
                                            key={pageNumber}
                                            onClick={() => setCurrentPage(pageNumber)}
                                            style={{
                                                cursor: 'pointer',
                                                padding: '6px 10px',
                                                borderRadius: '6px',
                                                fontWeight:
                                                    currentPage === pageNumber
                                                        ? 600
                                                        : 400,
                                                backgroundColor:
                                                    currentPage === pageNumber
                                                        ? '#E6F4F9'
                                                        : 'transparent',
                                                color:
                                                    currentPage === pageNumber
                                                        ? '#0A83AE'
                                                        : '#323130'
                                            }}
                                        >
                                            {pageNumber}
                                        </span>
                                    );
                                }
                            )}


                            <span
                                onClick={() => {
                                    if (currentPage < totalPages) {
                                        setCurrentPage(currentPage + 1);
                                    }
                                }}
                                style={{
                                    cursor:
                                        currentPage >= totalPages
                                            ? 'not-allowed'
                                            : 'pointer',
                                    opacity:
                                        currentPage >= totalPages
                                            ? 0.4
                                            : 1,
                                    padding: '6px',
                                    userSelect: 'none'
                                }}
                            >
                                ›
                            </span>

                            <span
                                onClick={() => setCurrentPage(totalPages)}
                                style={{
                                    cursor:
                                        currentPage >= totalPages
                                            ? 'not-allowed'
                                            : 'pointer',
                                    opacity:
                                        currentPage >= totalPages
                                            ? 0.4
                                            : 1,
                                    padding: '6px',
                                    userSelect: 'none'
                                }}
                            >
                                »
                            </span>
                        </div>
                    </div>
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
                                background: 'linear-gradient(135deg,#066C90,#0A83AE)',
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
                        backgroundColor: 'rgba(0,0,0,0.55)',
                        zIndex: 9999,
                        overflowY: 'auto',
                        padding: '24px 16px',
                        display: 'flex',
                        alignItems: 'flex-start',
                        justifyContent: 'center'
                    }}
                >
                    <div
                        style={{
                            width: '100%',
                            maxWidth: '1020px',
                            backgroundColor: '#fff',
                            borderRadius: '16px',
                            overflow: 'hidden',
                            boxShadow: '0 20px 60px rgba(0,0,0,0.25)',
                            margin: 'auto'
                        }}
                    >
                        {/* ── Header ── */}
                        <div
                            style={{
                                background: 'linear-gradient(135deg,#005f7a,#0A83AE)',
                                color: '#fff',
                                padding: '18px 24px 18px 28px',
                                position: 'relative',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '14px'
                            }}
                        >
                            <div
                                style={{
                                    width: '40px',
                                    height: '40px',
                                    borderRadius: '10px',
                                    backgroundColor: 'rgba(255,255,255,0.15)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    fontSize: '20px',
                                    flexShrink: 0
                                }}
                            >
                                {isEditMode ? '✏️' : '✨'}
                            </div>

                            <div style={{ flex: 1 }}>
                                <h2
                                    style={{
                                        margin: 0,
                                        fontSize: '20px',
                                        fontWeight: 700,
                                        letterSpacing: '-0.2px'
                                    }}
                                >
                                    {isEditMode ? 'Edit Prompt' : 'Add New Prompt'}
                                </h2>
                                <p
                                    style={{
                                        margin: '2px 0 0',
                                        fontSize: '13px',
                                        opacity: 0.85,
                                        fontWeight: 400
                                    }}
                                >
                                    {isEditMode
                                        ? 'Update this prompt and send it back for approval.'
                                        : 'Share a useful AI prompt with Stewart Title users.'}
                                </p>
                            </div>

                            <button
                                onClick={isSaving ? undefined : closeAddPromptModal}
                                disabled={isSaving}
                                style={{
                                    width: '36px',
                                    height: '36px',
                                    borderRadius: '8px',
                                    border: '1px solid rgba(255,255,255,0.25)',
                                    background: 'rgba(255,255,255,0.1)',
                                    color: '#fff',
                                    cursor: isSaving ? 'not-allowed' : 'pointer',
                                    fontSize: '18px',
                                    opacity: isSaving ? 0.4 : 1,
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    flexShrink: 0,
                                    lineHeight: 1
                                }}
                            >
                                ✕
                            </button>
                        </div>

                        {/* ── Body ── */}
                        <div
                            style={{
                                padding: '16px 20px 20px',
                                backgroundColor: '#f0f4f8',
                                position: 'relative'
                            }}
                        >
                            {/* Saving overlay */}
                            {isSaving && (
                                <div
                                    style={{
                                        position: 'absolute',
                                        inset: 0,
                                        backgroundColor: 'rgba(240,244,248,0.92)',
                                        zIndex: 10,
                                        display: 'flex',
                                        flexDirection: 'column',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        gap: '12px',
                                        borderRadius: '0 0 16px 16px',
                                        cursor: 'wait'
                                    }}
                                >
                                    <div
                                        style={{
                                            width: '40px',
                                            height: '40px',
                                            border: '3px solid rgba(10,131,174,0.2)',
                                            borderTopColor: '#0A83AE',
                                            borderRadius: '50%',
                                            animation: 'cpsp-spin 0.75s linear infinite'
                                        }}
                                    />
                                    <p
                                        style={{
                                            margin: 0,
                                            color: '#0A83AE',
                                            fontWeight: 600,
                                            fontSize: '14px'
                                        }}
                                    >
                                        {isEditMode ? 'Saving changes...' : 'Submitting prompt...'}
                                    </p>
                                </div>
                            )}

                            {/* ── Card: Prompt Details ── */}
                            <div
                                style={{
                                    backgroundColor: '#ffffff',
                                    borderRadius: '12px',
                                    padding: '18px 20px 20px',
                                    marginBottom: '12px',
                                    boxShadow: '0 1px 4px rgba(0,0,0,0.07)'
                                }}
                            >
                                {/* Section header */}
                                <div
                                    style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: '8px',
                                        marginBottom: '14px',
                                        paddingBottom: '10px',
                                        borderBottom: '1px solid #eef0f3'
                                    }}
                                >
                                    <span style={{ fontSize: '15px' }}>📋</span>
                                    <span
                                        style={{
                                            fontSize: '11px',
                                            fontWeight: 700,
                                            letterSpacing: '0.8px',
                                            color: '#006d8f',
                                            textTransform: 'uppercase'
                                        }}
                                    >
                                        Prompt Details
                                    </span>
                                </div>

                                {/* Row 1 – Prompt Name | Tags */}
                                <div
                                    style={{
                                        display: 'grid',
                                        gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
                                        columnGap: '16px',
                                        marginBottom: '14px'
                                    }}
                                >
                                    <div>
                                        <label
                                            style={{
                                                display: 'block',
                                                marginBottom: '5px',
                                                fontWeight: 600,
                                                fontSize: '13px',
                                                color: '#323130'
                                            }}
                                        >
                                            Prompt Name <span style={{ color: '#c50f1f' }}>*</span>
                                        </label>
                                        <input
                                            value={newPrompt.title}
                                            onChange={(e) =>
                                                setNewPrompt({
                                                    ...newPrompt,
                                                    title: e.target.value
                                                })
                                            }
                                            placeholder="e.g. Summarise meeting notes"
                                            style={{
                                                width: '100%',
                                                padding: '9px 12px',
                                                borderRadius: '8px',
                                                border: '1px solid #d1d5db',
                                                fontSize: '13px',
                                                boxSizing: 'border-box',
                                                color: '#1f2937',
                                                backgroundColor: '#fafafa',
                                                outline: 'none'
                                            }}
                                        />
                                    </div>

                                    <div>
                                        <label
                                            style={{
                                                display: 'block',
                                                marginBottom: '5px',
                                                fontWeight: 600,
                                                fontSize: '13px',
                                                color: '#323130'
                                            }}
                                        >
                                            Tags
                                            <span
                                                style={{
                                                    marginLeft: '6px',
                                                    fontWeight: 400,
                                                    color: '#8a8a8a',
                                                    fontSize: '12px'
                                                }}
                                            >
                                                (optional, comma-separated)
                                            </span>
                                        </label>
                                        <input
                                            value={newPrompt.tags}
                                            onChange={(e) =>
                                                setNewPrompt({
                                                    ...newPrompt,
                                                    tags: e.target.value
                                                })
                                            }
                                            placeholder="e.g. productivity, email, summarise"
                                            style={{
                                                width: '100%',
                                                padding: '9px 12px',
                                                borderRadius: '8px',
                                                border: '1px solid #d1d5db',
                                                fontSize: '13px',
                                                boxSizing: 'border-box',
                                                color: '#1f2937',
                                                backgroundColor: '#fafafa',
                                                outline: 'none'
                                            }}
                                        />
                                    </div>
                                </div>

                                {/* Row 2 – AI Tool | Department */}
                                <div
                                    style={{
                                        display: 'grid',
                                        gridTemplateColumns: 'repeat(2, minmax(0, 1fr))',
                                        columnGap: '16px',
                                        marginBottom: '14px'
                                    }}
                                >
                                    <div>
                                        <label
                                            style={{
                                                display: 'block',
                                                marginBottom: '5px',
                                                fontWeight: 600,
                                                fontSize: '13px',
                                                color: '#323130'
                                            }}
                                        >
                                            AI Tool <span style={{ color: '#c50f1f' }}>*</span>
                                        </label>
                                        <select
                                            value={newPrompt.aiTool}
                                            onChange={(e) =>
                                                setNewPrompt({
                                                    ...newPrompt,
                                                    aiTool: e.target.value
                                                })
                                            }
                                            style={{
                                                width: '100%',
                                                padding: '9px 12px',
                                                borderRadius: '8px',
                                                border: '1px solid #d1d5db',
                                                fontSize: '13px',
                                                boxSizing: 'border-box',
                                                color: newPrompt.aiTool ? '#1f2937' : '#8a8a8a',
                                                backgroundColor: '#fafafa',
                                                outline: 'none'
                                            }}
                                        >
                                            <option value="">Select AI Tool</option>
                                            {aiToolOptions.map((tool: string) => (
                                                <option key={tool} value={tool}>{tool}</option>
                                            ))}
                                        </select>
                                    </div>

                                    <div>
                                        <label
                                            style={{
                                                display: 'block',
                                                marginBottom: '5px',
                                                fontWeight: 600,
                                                fontSize: '13px',
                                                color: '#323130'
                                            }}
                                        >
                                            Department <span style={{ color: '#c50f1f' }}>*</span>
                                        </label>
                                        <select
                                            value={newPrompt.department}
                                            onChange={(e) =>
                                                setNewPrompt({
                                                    ...newPrompt,
                                                    department: e.target.value
                                                })
                                            }
                                            style={{
                                                width: '100%',
                                                padding: '9px 12px',
                                                borderRadius: '8px',
                                                border: '1px solid #d1d5db',
                                                fontSize: '13px',
                                                boxSizing: 'border-box',
                                                color: newPrompt.department ? '#1f2937' : '#8a8a8a',
                                                backgroundColor: '#fafafa',
                                                outline: 'none'
                                            }}
                                        >
                                            <option value="">Select Department</option>
                                            {departmentOptions.map((department: string) => (
                                                <option key={department} value={department}>{department}</option>
                                            ))}
                                        </select>
                                    </div>
                                </div>

                                {/* Row 3 – Use Case (full width) */}
                                <div>
                                    <label
                                        style={{
                                            display: 'block',
                                            marginBottom: '5px',
                                            fontWeight: 600,
                                            fontSize: '13px',
                                            color: '#323130'
                                        }}
                                    >
                                        Use Case <span style={{ color: '#c50f1f' }}>*</span>
                                    </label>
                                    <textarea
                                        rows={3}
                                        value={newPrompt.useCase}
                                        onChange={(e) =>
                                            setNewPrompt({
                                                ...newPrompt,
                                                useCase: e.target.value
                                            })
                                        }
                                        placeholder="Describe when and why this prompt is useful…"
                                        style={{
                                            width: '100%',
                                            padding: '9px 12px',
                                            borderRadius: '8px',
                                            border: '1px solid #d1d5db',
                                            fontSize: '13px',
                                            resize: 'vertical',
                                            boxSizing: 'border-box',
                                            color: '#1f2937',
                                            backgroundColor: '#fafafa',
                                            outline: 'none',
                                            fontFamily: 'inherit',
                                            lineHeight: '1.5'
                                        }}
                                    />
                                </div>
                            </div>

                            {/* ── Card: Prompt Text ── */}
                            <div
                                style={{
                                    backgroundColor: '#ffffff',
                                    borderRadius: '12px',
                                    padding: '18px 20px 20px',
                                    marginBottom: '14px',
                                    boxShadow: '0 1px 4px rgba(0,0,0,0.07)'
                                }}
                            >
                                {/* Section header */}
                                <div
                                    style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: '8px',
                                        marginBottom: '12px',
                                        paddingBottom: '10px',
                                        borderBottom: '1px solid #eef0f3'
                                    }}
                                >
                                    <span style={{ fontSize: '15px' }}>💬</span>
                                    <span
                                        style={{
                                            fontSize: '11px',
                                            fontWeight: 700,
                                            letterSpacing: '0.8px',
                                            color: '#006d8f',
                                            textTransform: 'uppercase'
                                        }}
                                    >
                                        Prompt Text
                                    </span>
                                    <span
                                        style={{
                                            marginLeft: 'auto',
                                            fontSize: '11px',
                                            color: '#8a8a8a',
                                            fontWeight: 400
                                        }}
                                    >
                                        The exact text that will be sent to the AI
                                    </span>
                                </div>

                                <textarea
                                    rows={7}
                                    value={newPrompt.promptText}
                                    onChange={(e) =>
                                        setNewPrompt({
                                            ...newPrompt,
                                            promptText: e.target.value
                                        })
                                    }
                                    placeholder="Enter the full prompt text here…"
                                    style={{
                                        width: '100%',
                                        padding: '12px 14px',
                                        borderRadius: '8px',
                                        border: '1px solid #d1d5db',
                                        fontSize: '13px',
                                        resize: 'vertical',
                                        boxSizing: 'border-box',
                                        color: '#1f2937',
                                        backgroundColor: '#fafafa',
                                        outline: 'none',
                                        fontFamily: 'inherit',
                                        lineHeight: '1.6'
                                    }}
                                />
                            </div>

                            {/* ── Footer ── */}
                            <div
                                style={{
                                    display: 'flex',
                                    justifyContent: 'flex-end',
                                    alignItems: 'center',
                                    gap: '10px',
                                    paddingTop: '4px'
                                }}
                            >
                                <button
                                    onClick={isSaving ? undefined : closeAddPromptModal}
                                    disabled={isSaving}
                                    style={{
                                        padding: '9px 20px',
                                        borderRadius: '8px',
                                        border: '1px solid #d1d5db',
                                        backgroundColor: '#ffffff',
                                        color: '#323130',
                                        fontSize: '13px',
                                        fontWeight: 500,
                                        cursor: isSaving ? 'not-allowed' : 'pointer',
                                        opacity: isSaving ? 0.5 : 1
                                    }}
                                >
                                    Cancel
                                </button>

                                <button
                                    onClick={isSaving ? undefined : savePrompt}
                                    disabled={isSaving}
                                    style={{
                                        backgroundColor: isSaving ? '#7bbfd6' : '#0A83AE',
                                        color: '#fff',
                                        border: 'none',
                                        padding: '9px 22px',
                                        borderRadius: '8px',
                                        cursor: isSaving ? 'not-allowed' : 'pointer',
                                        fontWeight: 600,
                                        fontSize: '13px',
                                        boxShadow: isSaving ? 'none' : '0 2px 8px rgba(10,131,174,0.35)',
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: '7px'
                                    }}
                                >
                                    {isSaving && (
                                        <span
                                            style={{
                                                width: '14px',
                                                height: '14px',
                                                border: '2px solid rgba(255,255,255,0.4)',
                                                borderTopColor: '#fff',
                                                borderRadius: '50%',
                                                display: 'inline-block',
                                                animation: 'cpsp-spin 0.75s linear infinite',
                                                flexShrink: 0
                                            }}
                                        />
                                    )}
                                    {isSaving
                                        ? (isEditMode ? 'Saving...' : 'Submitting...')
                                        : (isEditMode ? 'Update and Send for Approval' : 'Send for Approval')
                                    }
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
            {/* Toast notification stack */}
            {toasts.length > 0 && (
                <div
                    style={{
                        position: 'fixed',
                        top: '24px',
                        right: '24px',
                        zIndex: 99999,
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '10px',
                        maxWidth: '400px',
                        minWidth: '300px'
                    }}
                >
                    {toasts.map(toast => (
                        <div
                            key={toast.id}
                            style={{
                                backgroundColor: '#ffffff',
                                borderRadius: '10px',
                                boxShadow: '0 4px 20px rgba(0,0,0,0.14)',
                                borderLeft: `5px solid ${
                                    toast.type === 'success' ? '#107C10' : '#D13438'
                                }`,
                                padding: '14px 16px',
                                display: 'flex',
                                alignItems: 'flex-start',
                                gap: '12px',
                                animation: 'cpsp-fadeIn 0.25s ease'
                            }}
                        >
                            <span
                                style={{
                                    fontSize: '18px',
                                    flexShrink: 0,
                                    lineHeight: '1.4'
                                }}
                            >
                                {toast.type === 'success' ? '✅' : '❌'}
                            </span>

                            <span
                                style={{
                                    flex: 1,
                                    fontSize: '14px',
                                    color: '#323130',
                                    lineHeight: '1.5'
                                }}
                            >
                                {toast.message}
                            </span>

                            <button
                                onClick={() => dismissToast(toast.id)}
                                style={{
                                    background: 'none',
                                    border: 'none',
                                    cursor: 'pointer',
                                    color: '#605e5c',
                                    fontSize: '16px',
                                    padding: '0',
                                    flexShrink: 0,
                                    lineHeight: 1
                                }}
                            >
                                ✕
                            </button>
                        </div>
                    ))}
                </div>
            )}
        </div >
    );
};

export default Dashboard;