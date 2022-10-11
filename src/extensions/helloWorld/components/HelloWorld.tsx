/* eslint-disable no-useless-concat */
/* eslint-disable @typescript-eslint/no-unused-vars */
import * as React from 'react';
import { Log, FormDisplayMode, Guid } from '@microsoft/sp-core-library';
//import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './HelloWorld.module.scss';
import { SPFI } from '@pnp/sp';
import { useEffect, useState } from 'react';
import NewForm from '../../../customComponents/NewForm';
import DisplayForm from '../../../customComponents/DisplayForm';
import EditForm from '../../../customComponents/EditForm';
import { IRelatedItem } from '../../../models/interfaces/IRelatedItem';

//import {sp} from '@pnp/sp/presets/all';
import { Web } from "@pnp/sp/webs";
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';


export interface IHelloWorldProps {
  context: FormCustomizerContext;
  sp:SPFI;
  listGuid:Guid;
  itemID:number;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}



const LOG_SOURCE: string = 'HelloWorld';

const TestCustomizer: React.FC<IHelloWorldProps> = (props) => {
    useEffect(() => {
        Log.info(LOG_SOURCE, 'React Element: TestCustomizer mounted');
        console.log("I mounted")
        return () => {
            Log.info(LOG_SOURCE, 'React Element: TestCustomizer unmounted');
        }
    }, []);
    

//     //BYT SÅ DET INTE BLIR HÅRDKODAT SEDAN
// const web = Web("https://wcqvp.sharepoint.com/sites/sparvagen/");


//   const [relatedItems, setRelatedItems] = useState();

//    async function getRelatedItems():Promise<IRelatedItem[]>{
//     console.log("I get related items")
//     try {
//         const relatedItems = await web.lists.getByTitle("Vningar")
//         .items
//         .select("Title")
//         ();
//         return relatedItems as unknown as Promise<IRelatedItem[]>;
//       } catch (error) {
//         console.log("ERROR")
//         console.error(error);
//       }
//   }



  return (
  <div className={styles.helloWorld}>
        <h1>Felanmälan</h1>      
    {props.displayMode === FormDisplayMode.New &&
      <NewForm context={props.context} sp={props.sp} listGuid={props.listGuid} onSave={props.onSave}
          onClose={props.onClose} />
      }
      {props.displayMode === FormDisplayMode.Edit &&
          <EditForm sp={props.sp} listGuid={props.listGuid} itemId={props.itemID}
              onSave={props.onSave} onClose={props.onClose} />
      }
      {props.displayMode === FormDisplayMode.Display &&
          <DisplayForm sp={props.sp} listGuid={props.listGuid} itemId={props.itemID}
              onClose={props.onClose} />
      }
  </div>);
};

export default TestCustomizer;
