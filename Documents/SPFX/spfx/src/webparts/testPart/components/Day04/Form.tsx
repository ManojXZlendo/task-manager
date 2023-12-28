/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import {
    TextField,
    PrimaryButton,
    List,
    IconButton,
    MessageBar,
    MessageBarType,
    Spinner,
    DefaultButton
} from '@fluentui/react';

import {
    SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions
} from '@microsoft/sp-http';

interface IFormProps {
    context: {
        spHttpClient: SPHttpClient;
        pageContext: {
            web: {
                absoluteUrl: string;
            };
        };
    };
}

interface IMyFormState {
    Title: string;
    Email: string;
    editId: number | null;
    items: any;
    toastMessage: string | null;
    toastType: 'success' | 'error' | null;
    loading: boolean;
}

export class Form extends React.Component<IFormProps, IMyFormState> {
    constructor(props: IFormProps) {
        super(props);
        this.state = {
            Title: '',
            Email: '',
            editId: null,
            items: [],
            toastMessage: null,
            toastType: null,
            loading: false
        };
    }

    componentDidMount() {
        this.loadItems();
    }
    private displayToast(message: string, type: 'success' | 'error' = 'success', timeout: number = 2000): void {
        this.setState({ toastMessage: message, toastType: type });
        setTimeout(() => {
            this.setState({ toastMessage: null, toastType: null });
        }, timeout);
    }
    private async loadItems(): Promise<void> {
        this.setState({ loading: true });
        try {
            const endpoint = `${this.props.context.pageContext.web.absoluteUrl}/sites/SPTraining/_api/web/lists/getbytitle('SP_Training_Form')/items`;
            const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
            const items = await response.json();
            this.setState({ items });
            this.setState({ loading: false });
        }
        catch {
            this.displayToast('Error while loading items', 'error');
            this.setState({ items: [] });
            this.setState({ loading: false });
        }
        finally {
            this.setState({ loading: false });
        }
    }

    private async createItem(title: string, author: string): Promise<void> {
        const endpoint: string = `${this.props.context.pageContext.web.absoluteUrl}/sites/SPTraining/_api/web/lists/getbytitle('SP_Training_Form')/items`;
        const itemBody: any = {
            Title: title,
            Email: author
        };
        const requestOptions: ISPHttpClientOptions = {
            body: JSON.stringify(itemBody)
        };
        await this.props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, requestOptions);
        this.displayToast('Added successfully!');
        this.loadItems();
    }

    private async updateItem(id: number, title: string, author: string): Promise<void> {
        const endpoint: string = `${this.props.context.pageContext.web.absoluteUrl}/sites/SPTraining/_api/web/lists/getbytitle('SP_Training_Form')/items(${id})`;
        const itemBody: any = {
            Title: title,
            Email: author
        };
        const requestOptions: ISPHttpClientOptions = {
            headers: {
                'X-HTTP-Method': 'MERGE',
                'IF-MATCH': '*'
            },
            body: JSON.stringify(itemBody)
        };
        await this.props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, requestOptions);
        this.displayToast('Updated successfully!');
        this.loadItems();
    }

    private async deleteItem(id: number): Promise<void> {
        const endpoint: string = `${this.props.context.pageContext.web.absoluteUrl}/sites/SPTraining/_api/web/lists/getbytitle('SP_Training_Form')/items(${id})`;
        const requestOptions: ISPHttpClientOptions = {
            headers: {
                'X-HTTP-Method': 'DELETE',
                'IF-MATCH': '*'
            }
        };
        await this.props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, requestOptions);
        this.displayToast('Deleted successfully!');
        this.loadItems();
    }

    private async handleSubmit(event: React.FormEvent<HTMLFormElement>): Promise<void> {
        event.preventDefault();
        const { Title, Email, editId } = this.state;
        if (editId) {
            await this.updateItem(editId, Title, Email);
        } else {
            await this.createItem(Title, Email);
        }
        this.setState({ Title: '', Email: '', editId: null });
    }

    render() {
        const { Title, Email, items, toastMessage, toastType, loading } = this.state;
        const ListOfData = items?.value;

        return (
            <div style={{ padding: '20px' }}>
                {loading && <Spinner label="Loading..." />}
                {!loading && <form onSubmit={this.handleSubmit.bind(this)}>
                    <TextField
                        label="Title"
                        value={Title}
                        onChange={(e, newValue) => this.setState({ Title: newValue || '' })}
                        placeholder="Title"
                        required
                    />
                    <TextField
                        label="Email "
                        value={Email}
                        onChange={(e, newValue) => this.setState({ Email: newValue || '' })}
                        placeholder="Enter Email Address"
                        required
                    />
                    <PrimaryButton disabled={!Title || !Email} type="submit" style={{ marginTop: '10px' }}>
                        {this.state.editId ? 'Update' : 'Add'}
                    </PrimaryButton>
                    {this.state.editId &&
                        <DefaultButton style={{ marginLeft: '10px' }}
                            onClick={() => this.setState({ editId: null, Title: '', Email: '' })}>Cancel</DefaultButton>}
                </form>}

                {!loading && !this.state.editId && ListOfData &&
                    <List items={ListOfData} onRenderCell={(item, index) => (
                        <div
                            key={item.ID}
                            style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0' }}>
                            <span>{item.Title} by {item.Email}</span>
                            <div>
                                <IconButton
                                    iconProps={{ iconName: 'Edit' }}
                                    title="Edit"
                                    ariaLabel="Edit"
                                    onClick={() => this.setState({
                                        Title: item.Title,
                                        Email: item.Email,
                                        editId: item.ID
                                    })}
                                />
                                <IconButton
                                    iconProps={{ iconName: 'Delete' }}
                                    title="Delete"
                                    ariaLabel="Delete"
                                    onClick={() => this.deleteItem(item.ID)}
                                />
                            </div>
                        </div>
                    )} />}

                {!loading && toastMessage && toastType && (
                    <MessageBar
                        messageBarType={toastType === 'success' ? MessageBarType.success : MessageBarType.error}
                        isMultiline={false}

                        onDismiss={() => this.setState({ toastMessage: null, toastType: null })}
                        dismissButtonAriaLabel="Close"
                    >
                        {toastMessage}
                    </MessageBar>
                )}
            </div>
        );
    }
}
