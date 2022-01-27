import { Log } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";

const LOG_SOURCE: string = "My Teams Graph";
const PROFILE_IMAGE_URL: string =
    "/_layouts/15/userphoto.aspx?size=M&accountname=";
const DEFAULT_PERSONA_IMG_HASH: string = "7ad602295f8386b7615b582d87bcc294";
const DEFAULT_IMAGE_PLACEHOLDER_HASH: string =
    "4a48f26592f4e1498d7a478a4c48609c";
const MD5_MODULE_ID: string = "8494e7d7-6b99-47b2-a741-59873e42f16f";

export class SharePointServiceProvider {
    constructor(private _context: WebPartContext) {
        // Setup Context to PnPjs and MSGraph
        sp.setup({
            spfxContext: this._context
        });

        this.onInit();
    }

    private async onInit() { }

    public async getUserPhoto(userId): Promise<string> {
        const personaImgUrl = PROFILE_IMAGE_URL + userId;
        const url: string = await this.getImageBase64(personaImgUrl);
        const newHash = await this.getMd5HashForUrl(url);

        if (
            newHash !== DEFAULT_PERSONA_IMG_HASH &&
            newHash !== DEFAULT_IMAGE_PLACEHOLDER_HASH
        ) {
            return "data:image/png;base64," + url;
        } else {
            return "undefined";
        }
    }

    /**
   * Gets image base64
   * @param pictureUrl
   * @returns image base64
   */
    private getImageBase64(pictureUrl: string): Promise<string> {
        return new Promise((resolve, reject) => {
            let image = new Image();
            image.addEventListener("load", () => {
                let tempCanvas = document.createElement("canvas");
                (tempCanvas.width = image.width),
                    (tempCanvas.height = image.height),
                    tempCanvas.getContext("2d").drawImage(image, 0, 0);
                let base64Str;
                try {
                    base64Str = tempCanvas.toDataURL("image/png");
                } catch (e) {
                    return "";
                }
                base64Str = base64Str.replace(/^data:image\/png;base64,/, "");
                resolve(base64Str);
            });
            image.src = pictureUrl;
        });
    }

    /**
     * Get MD5Hash for the image url to verify whether user has default image or custom image
     * @param url
     */
    private getMd5HashForUrl(url: string) {
        return new Promise(async (resolve, reject) => {
            const library: any = await this.loadSPComponentById(MD5_MODULE_ID);
            try {
                const md5Hash = library.Md5Hash;
                if (md5Hash) {
                    const convertedHash = md5Hash(url);
                    resolve(convertedHash);
                }
            } catch (error) {
                resolve(url);
            }
        });
    }

    /**
     * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
     * @param componentId - componentId, guid of the component library
     */
    private loadSPComponentById(componentId: string) {
        return new Promise((resolve, reject) => {
            SPComponentLoader.loadComponentById(componentId)
                .then((component: any) => {
                    resolve(component);
                })
                .catch(error => { });
        });
    }

    /**
     * Gets user profile
     * @param loginName
     * @returns user profile
     */
    public async getUserProfile(loginName: string): Promise<any> {
        try {
            const _loginName = `i:0#.f|membership|${loginName}`;
            const user = await sp.profiles.getPropertiesFor(_loginName);
            return user;
        } catch (error) {
            Log.error(LOG_SOURCE, error, this._context.serviceScope);
            throw new Error(error.message);
        }
    }
}