import { OfficeState } from '../reducers/Office'

export class OfficeService {
    promiseInitialize(): Promise<OfficeState> {
        let resolver: (value?: OfficeState | PromiseLike<OfficeState>) => void;
        let promise = new Promise<OfficeState>((resolve, reject) => {
            resolver = resolve
        });
        Office.initialize = function (reason) {
            resolver();
        }
        return promise;
    }
}