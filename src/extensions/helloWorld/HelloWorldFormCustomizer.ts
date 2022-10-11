import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import HelloWorld, { IHelloWorldProps } from './components/HelloWorld';
import { SPFI, spfi, SPFx } from '@pnp/sp';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = 'HelloWorldFormCustomizer';

export default class HelloWorldFormCustomizer
extends BaseFormCustomizer<IHelloWorldFormCustomizerProperties> {
  
  private  sp: SPFI = undefined;

  public onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated HelloWorldFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    this.sp =spfi().using(SPFx(this.context));
    return Promise.resolve();
  }

  public render(): void {
    // Use this method to perform your custom rendering.

    const helloWorld: React.ReactElement<{}> =
      React.createElement(HelloWorld, {
        sp:this.sp,
        context: this.context,
        listGuid:this.context.list.guid,
        itemID:this.context.itemId,
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose
       } as  unknown as IHelloWorldProps);

    ReactDOM.render(helloWorld, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {

    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
