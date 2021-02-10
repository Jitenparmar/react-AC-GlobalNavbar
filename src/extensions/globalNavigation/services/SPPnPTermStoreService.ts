
import { sp } from "@pnp/sp";
import "@pnp/sp/taxonomy";
import { ITermInfo } from "@pnp/sp/taxonomy";
import { dateAdd, PnPClientStorage } from "@pnp/common";

const LOG_SOURCE: string = 'SPPnPTermStoreService';


/**
 * @interface
 * Interface for SPTermStoreService configuration
 */



export class SPPnPTermStoreService{

    /**
     * @function
     * Service constructor
     */
    constructor() {

    }

     /**
     * @function
     * Gets the collection of term stores in the current SharePoint env
     */

    public async getTermsFromTermSetAsync(termSetId: string, termSetLocal: Number): Promise<any[]> {
       // here we get all the children of a given set
        const childTree = await sp.termStore.groups.getById(termSetId).sets.getById(termSetId).getAllChildrenAsOrderedTree();
        
        // here we show caching the results using the PnPClientStorage class, there are many caching libraries and options available
        const store = new PnPClientStorage();
        
        // our tree likely doesn't change much in 30 minutes for most applications
        // adjust to be longer or shorter as needed
        const cachedTree = await store.local.getOrPut("myKey", () => {
            return sp.termStore.groups.getById(termSetId).sets.getById(termSetId).getAllChildrenAsOrderedTree();
        }, dateAdd(new Date(), "minute", 30));

        return childTree;
    }

}