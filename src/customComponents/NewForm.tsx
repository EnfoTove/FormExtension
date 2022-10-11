/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, FC, useEffect } from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { SPFI, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
import styles from './NewForm.module.scss';
import { IRelatedItem } from '../models/interfaces/IRelatedItem';
import { Web } from '@pnp/sp/webs';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { Dropdown, IDropdownOption, IDropdownStyles, IStackTokens, Stack } from 'office-ui-fabric-react';
 
export interface INewFormProps {
    context:FormCustomizerContext;
    sp: SPFI;
    listGuid: Guid;
    onSave: () => void;
    onClose: () => void;
}
//let floorItems;
const NewForm: FC<INewFormProps> = (props) => {
    const [buildingTitle, setBuildingTitle] = useState<string>('');
    const [floorArray, setfloorArray] =  useState<IRelatedItem[]>([]);
    const [errorDescription, setErrorDescription] = useState<any>(undefined);
    const [msg, setMsg] = useState<any>(undefined);
    const [optionsLoaded, setOptionsLoaded] = useState(false);
    const options: IDropdownOption[]=[]

    function sleep(ms:number) {
        return new Promise(resolve => setTimeout(resolve, ms));
      }

     useEffect(() => {
        const handleInputChange = async () => {
            await sleep(1000)
            .then(populateFloorArray)
            //.then(createOptions);
        }
        handleInputChange()
        .catch(console.error);
      }, [buildingTitle]);
 
    const clearControls = () => {
        setBuildingTitle('');
        setErrorDescription('');
    };


    const dropdownStyles: Partial<IDropdownStyles> = {
        dropdown: { width: 300 },
      };
 
    const saveListItem = async () => {
        setMsg(undefined);
        await props.sp.web.lists.getById(props.listGuid.toString()).items.add({
            Title: buildingTitle,
            Felbeskrivning:errorDescription
        });
        setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
        clearControls();
    };

        //BYT SÅ DET INTE BLIR HÅRDKODAT SEDAN
    const web = Web("https://wcqvp.sharepoint.com/sites/sparvagen/").using(SPFx(props.context));

 async function getRelatedItems():Promise<IRelatedItem[]>{
     const queryString = `Byggnad eq  '${ buildingTitle }'`;
  try {
      const relatedItems = await web.lists.getByTitle("Våningar")
      .items
      .select("Title, Byggnad")
      .filter(queryString)
      ();
      return relatedItems as unknown as Promise<IRelatedItem[]>;
    } catch (error) {
      console.error(error);
    }
}    
    const stackTokens: IStackTokens = { childrenGap: 20 };

    async function populateFloorArray(items:any) {
        try {
            await getRelatedItems().then(items=>{
                setfloorArray(items)
            })
        } catch (error) {
            console.error(error);
        }
    }

    
    
    const createOptions=()=>{
        console.log("Options loading")
        floorArray.forEach(floor => {
            const object= { key: floor.Title, text: floor.Title};        
                options.push(object);              
            });               
      //  setOptionsLoaded(true)
    }
    
   

    return (
        <div className={styles.newForm}>
            <div className={styles.newFormInput}>
                <TextField label="Ange Fastighetsnamn" value={buildingTitle} onChange={(e,v) => setBuildingTitle(v)} />
            </div>
            <Stack tokens={stackTokens}>
                <Dropdown
                    placeholder="Våning"
                    label="Välj våning"
                    options={options}
                    styles={dropdownStyles}
                    onClick={createOptions}
                />
            </Stack>
            <div className={styles.newFormInput}>
                <TextField label="Ange felbeskrivning" value={errorDescription} onChange={(e, v) => setErrorDescription(v)} />
            </div>
                <PrimaryButton text="Save" onClick={saveListItem} />
            {msg && msg.Message &&
                <MessageBar messageBarType={msg.scope ? msg.scope : MessageBarType.info}>{msg.Message}</MessageBar>
            }
        </div>
    );
};
 
export default NewForm;