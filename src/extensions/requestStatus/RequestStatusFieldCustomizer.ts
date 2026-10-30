import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'RequestStatusFieldCustomizerStrings';
import styles from './RequestStatusFieldCustomizer.module.scss';

export interface IRequestStatusFieldCustomizerProperties {
  sampleText?: string;
}

const LOG_SOURCE: string = 'RequestStatusFieldCustomizer';

export default class RequestStatusFieldCustomizer
  extends BaseFieldCustomizer<IRequestStatusFieldCustomizerProperties> {
  
  public onInit(): Promise<void> {
    console.log('[RequestStatus] onInit executed');
    Log.info(LOG_SOURCE, 'RequestStatusFieldCustomizer initialized');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(
      LOG_SOURCE,
      `Extension title: ${strings.Title}`
    );

    return Promise.resolve();
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const rawValue: unknown = event.fieldValue;
    const status = String(rawValue ?? '').trim();
    const normalizedStatus = status.toLowerCase();
    console.log('[RequestStatus] onRenderCell', rawValue);

    const statusClass =
      normalizedStatus === 'approved'
        ? styles.approved
        : normalizedStatus === 'rejected'
          ? styles.rejected
          : normalizedStatus === 'pending'
            ? styles.pending
            : styles.unknown;

    const label = status || 'Not set';

    event.domElement.replaceChildren();

    const pill = document.createElement('span');
    pill.classList.add(styles.requestStatus, statusClass);
    pill.textContent = label;

    event.domElement.appendChild(pill);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    super.onDisposeCell(event);
  }
}
