/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from 'react';
import { TextField, PrimaryButton, Stack, DefaultButton, Dialog, DialogType, DialogFooter } from '@fluentui/react';

interface IFormState {
    name: string;
    email: string;
    message: string;
    isDialogVisible: boolean;
}

const formStyle = {
    width: '100%',
    '@media (min-width: 768px)': {
        maxWidth: '600px' // Tablet view
    },
    '@media (min-width: 320px)': {
        maxWidth: '300px' // Mobile view
    }
};

class Form extends React.Component<{}, IFormState> {
    constructor(props: {}) {
        super(props);
        this.state = {
            name: '',
            email: '',
            message: '',
            isDialogVisible: false
        };
    }

    handleChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
        const { name } = event.currentTarget;
        this.setState({
            [name]: newValue
        } as unknown as Pick<IFormState, keyof IFormState>);
    };

    handleSubmit = (event: React.FormEvent) => {
        event.preventDefault();
        //ignore isDialogVisible
        const { isDialogVisible, ...formData } = this.state;
        console.log(formData); // Handle form submission logic here
    };


    handleClear = () => {
        this.setState({
            name: '',
            email: '',
            message: '',
            isDialogVisible: true
        });
    };

    handleDialogClose = () => {
        this.setState({ isDialogVisible: false });
    };

    render() {
        return (
            <Stack horizontalAlign="center" verticalAlign="start" tokens={{ childrenGap: 15, padding: '20px' }}>
                <form onSubmit={this.handleSubmit} style={formStyle}>
                    <Stack horizontal wrap tokens={{ childrenGap: 15 }} styles={{ root: { width: '100%' } }}>
                        <TextField
                            label="Name"
                            name="name"
                            value={this.state.name}
                            onChange={this.handleChange}
                            required
                            styles={{ root: { width: '100%' } }}
                        />
                        <TextField
                            label="Email"
                            name="email"
                            value={this.state.email}
                            onChange={this.handleChange}
                            required
                            type="email"
                            styles={{ root: { width: '100%' } }}
                        />
                        <TextField
                            label="Message"
                            name="message"
                            value={this.state.message}
                            onChange={this.handleChange}
                            multiline
                            rows={4}
                            styles={{ root: { width: '100%' } }}
                        />
                        <PrimaryButton type="submit" styles={{ root: { width: '100%' } }}>Submit</PrimaryButton>
                        <DefaultButton text="Clear" onClick={this.handleClear} styles={{ root: { width: '100%' } }} />
                    </Stack>
                </form>

                <Dialog
                    hidden={!this.state.isDialogVisible}
                    onDismiss={this.handleDialogClose}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Clear Form',
                        subText: 'Are you sure you want to clear the form?'
                    }}
                    modalProps={{
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } }
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton onClick={this.handleDialogClose} text="Cancel" />
                        <DefaultButton onClick={this.handleDialogClose} text="Clear" />
                    </DialogFooter>
                </Dialog>
            </Stack>
        );
    }
}

export default Form;
