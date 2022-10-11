/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, FC, useEffect } from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { SPFI, SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Guid } from '@microsoft/sp-core-library';
import styles from './NewForm.module.scss';
import { IRelatedItem } from '../models/interfaces/IRelatedItem';
import { Web } from '@pnp/sp/webs';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { Dropdown, Icon, IDropdownOption, IDropdownStyles, IStackTokens, Stack } from 'office-ui-fabric-react';
import { IItem } from "@pnp/sp/items/types";
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
    const [selectedFile, setSelectedFile] = useState();
    const [selectedFileName, setSelectedFileName] = useState();
    const [msg, setMsg] = useState<any>(undefined);
    const options: IDropdownOption[]=[]
    let counter = 0;
    const AttachmentIcon = () => <Icon iconName="Attach" className={styles.attachmentIcon}/>;

            //BYT SÅ DET INTE BLIR HÅRDKODAT SEDAN
        //const web = Web("https://wcqvp.sharepoint.com/sites/sparvagen/").using(SPFx(props.context));
    

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
 
    const addListItem = async () => {
        let itemId:number;
        setMsg(undefined);
        await props.sp.web.lists.getById(props.listGuid.toString()).items.add({
            Title: buildingTitle,
            Felbeskrivning:errorDescription
        }).then((result)=>{
            itemId=result.data.ID;
            console.log(itemId)
        });

        console.log(props.listGuid.toString())
        const item: IItem  = props.sp.web.lists.getById(props.listGuid.toString()).items.getById(itemId);
        console.log(item)
        console.log(selectedFileName)
        await item.attachmentFiles.add(selectedFileName, "This is a message!!")
        
        setMsg({ scope: MessageBarType.success, Message: 'New item created successfully!' });
        clearControls();
    };

    // export async function PostMessage (itemObj:IInputItem, postList:string){
    //     let itemId:number;
    //     if(itemObj !== null){
    //     const messageItem=await sp.web.lists.getByTitle(postList).items.add({
    //       Title: itemObj.Rubrik,
    //       Meddelande: itemObj.Meddelande,
    //     }).then((result)=>{
    //       itemId=result.data.ID
    //     })
    //     const item: IItem  = sp.web.lists.getByTitle(postList).items.getById(itemId);
    //     await item.attachmentFiles.add(itemObj.BildNamn, itemObj.BildBlob)
    //   }
      
    //   }

 async function getRelatedItems():Promise<IRelatedItem[]>{
     const queryString = `Byggnad eq  '${ buildingTitle }'`;
  try {
      const relatedItems = await props.sp.web.lists.getByTitle("Våningar")
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
    
    const setCounter=()=>{
        // eslint-disable-next-line no-unused-expressions
        counter>0? null : createOptions()
        counter= counter +1;
    }

    const handleInputAndResetCounter=(value:string)=>{
        setBuildingTitle(value);
        counter=0;
    }

    const onFileChange = (event:any) => {
        setSelectedFileName(event.target.files[0].name)
        const file = event.target.files[0];
        const reader = new FileReader();
        // eslint-disable-next-line @typescript-eslint/no-empty-function
        reader.onloadend = function(){};
        reader.readAsDataURL(file);
        setSelectedFile(file);
    };
   

    return (
        <div className={styles.newForm}>
            <div className={styles.newFormInput}>
                <TextField label="Ange Fastighetsnamn" value={buildingTitle} onChange={(e,v) => handleInputAndResetCounter(v)} />
            </div>
            <Stack tokens={stackTokens}>
                <Dropdown
                    placeholder="Våning"
                    label="Välj våning"
                    options={options}
                    styles={dropdownStyles}
                    onClick={setCounter}
                />
            </Stack>
            <div className={styles.newFormInput}>
                <TextField label="Ange felbeskrivning" value={errorDescription} onChange={(e, v) => setErrorDescription(v)} />
            </div>
            <label className={styles.fileUploadLabel}>
            <input type="file" className={styles.defaultFileUploader} aria-label="File browser" onChange={onFileChange}/>
              <span className={styles.attachmentIconSpan}>
                <AttachmentIcon/>
                <span className={styles.selectedFile}>{selectedFileName}</span>
              </span>
          </label>
                <PrimaryButton text="Save" onClick={addListItem} />
            {msg && msg.Message &&
                <MessageBar messageBarType={msg.scope ? msg.scope : MessageBarType.info}>{msg.Message}</MessageBar>
            }
        </div>
    );
};
 
export default NewForm;