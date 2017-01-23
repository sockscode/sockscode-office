import { OFFICE_INITIALIZED, OfficeInitStateAction } from '../actions/Actions';

export interface OfficeState {
    initialized: boolean,
    supported: boolean,
    error?: string
}

const reducer = (state: OfficeState = { initialized: false, supported: false }, action: OfficeInitStateAction): OfficeState => {
    console.log(action);
    switch (action.type) {
        case OFFICE_INITIALIZED: {
            const {initialized, supported, error} = action;
            return { initialized, supported, error };
        }
        default:
            return state
    }
}

export const office = reducer;