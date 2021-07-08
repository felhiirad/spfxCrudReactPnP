import {sp} from '@pnp/sp/presets/all';
import { IDropdownOption } from 'office-ui-fabric-react';

export class SPOperations {
    public getListeTitles():Promise<IDropdownOption[]>{
        let listTitles:IDropdownOption[]=[];
        return new Promise<IDropdownOption[]>(async(resolve,reject)=>{
            sp.web.lists.select("Title")().then((results:any)=>{
                results.map((result)=>{
                    console.log(results)
                    listTitles.push({key:result.Title,text:result.Title})

                })
                resolve(listTitles)

            },(error:any)=>{reject("error accured")})

            
        })

    }
    //Create LIste Items
    public CreateListeItem(listTitles:string):Promise<string>{
        return new Promise<string>(async(resolve,reject)=>{
            sp.web.lists.getByTitle(listTitles).items.add({Title:"PnPJs Item"}).then((results:any)=>{
                resolve(" item with id " +results.data.ID+ " added succeflly")
            })
        })
}
//updated liste item
public UpdateListeItem(listTitle:string):Promise<string>{
    return new Promise<string>(async(resolve,reject)=>{
        this.getLatestItemId(listTitle)
        .then((itemId:number)=>{
            sp.web.lists
            .getByTitle(listTitle)
            .items
            .getById(itemId)
            .update({Title:" update PNP js item"})
            .then(()=>
                resolve("item with ID  " + itemId +" updated successffuly")
            )

        })
    })
}

//delete liste item
public DeleteListeItem(listTitle:string):Promise<string>{
    return new Promise<string>(async(resolve,reject)=>{
        this.getLatestItemId(listTitle)
        .then((itemId:number)=>{
            sp.web.lists
            .getByTitle(listTitle)
            .items
            .getById(itemId)
            .delete()
            .then(()=>
                resolve("item with ID " + itemId +" deleted succesfly")
            )

        })
    })
}

//get latest item id
public getLatestItemId(listTitle:string):Promise<number>{
    return new Promise<number>(async(resolve,reject)=>{
        sp.web.lists
        .getByTitle(listTitle)
        .items
        .select("ID")
        .orderBy("ID",false)
        .top(1)()
        .then((result:any)=>{
            resolve(result[0].ID)
        })
    })
}


}