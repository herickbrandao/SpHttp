/** SPHTTP 1.0.3 - TypeScript Version */

interface ISpHttpConfig {
    ID?: number | string;
    Id?: number | string;
    top?: number;
    select?: string | string[];
    expand?: string | string[];
    filter?: string;
    orderby?: string;
    avoidcache?: boolean;
    versions?: boolean;
    recursive?: boolean;
    library?: string;
    target?: string;
    delete?: string | number;
    recycle?: boolean;
    name?: string;
    startswith?: boolean;
    tname?: string;
    url?: string;
    action?: string;
    item?: any;
}

class SpHttp {
    public baseURL: string = "../";
    public cleanResponse: boolean = true;
    public digest: string | null = null;
    public headers: any = {
        "Accept": "application/json; odata=verbose"
    };
    private _headers: any = {};
    public timeout: number = 15000;
    public top: number = 5000;
    public version: string = "1.0.3";

    private _error(e: any, obj: any = {}): any {
        console.error(e, obj);
        return e;
    }

    public async items(list: string, obj: ISpHttpConfig = {}): Promise<any> {
        if (!list) {
            return this._error("Your object has not a list name", obj);
        }

        const ID = obj.ID || obj.Id || false;
        const top = obj.top || this.top || 5000;
        let requestUrl = "_api/lists/getbytitle('" + list + "')/items";
        let concatenation = "?";

        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        if (!obj.select) obj.select = [];
        if (!obj.expand) obj.expand = [];
        
        const selectStr = typeof obj.select === "object" ? obj.select.join() : obj.select;
        const expandStr = typeof obj.expand === "object" ? obj.expand.join() : obj.expand;

        if (ID) {
            requestUrl += "(" + ID + ")";
            if (obj.versions) requestUrl += "/versions";
            if (selectStr?.length) {
                requestUrl += concatenation + "$select=" + selectStr;
                concatenation = "&";
            }
            if (expandStr?.length) {
                requestUrl += concatenation + "$expand=" + expandStr;
                concatenation = "&";
            }
            if (obj.avoidcache) {
                requestUrl += concatenation + "$v=" + (window as any).crypto.randomUUID();
            }
            return this._rest(requestUrl, {}, { select: selectStr.split(",") });
        }

        requestUrl += concatenation + "$top=" + top;
        concatenation = "&";

        if (selectStr?.length) requestUrl += concatenation + "$select=" + selectStr;
        if (expandStr?.length) requestUrl += concatenation + "$expand=" + expandStr;
        if (obj.filter?.length) requestUrl += concatenation + "$filter=" + obj.filter;
        if (obj.avoidcache) requestUrl += concatenation + "$v=" + (window as any).crypto.randomUUID();

        if (!obj.recursive) {
            if (obj.orderby?.length) requestUrl += concatenation + "$orderby=" + obj.orderby;
            return this._rest(requestUrl, {}, { select: selectStr.split(",") });
        }

        return this._recursive(this.baseURL + requestUrl, { ...obj, select: selectStr });
    }

    public async add(list: string, obj: any): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest
        }, this.headers);

        const url = "_api/lists/getbytitle('" + list + "')/items";
        return this._rest(url, {
            method: "POST",
            body: typeof obj === "object" ? JSON.stringify(obj) : obj
        });
    }

    public async update(list: string, obj: any): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = obj.ID || obj.Id || null;

        if (typeof obj === 'object' && ID) {
            url += "(" + ID + ")";
        } else {
            throw ('[ERROR] Put request data is wrong!');
        }

        return this._rest(url, {
            method: "POST",
            body: typeof obj === 'object' ? JSON.stringify(obj) : obj
        });
    }

    public async delete(list: string, obj: any): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = typeof obj === 'object' ? (obj.ID || obj.Id) : (typeof obj === 'number' ? obj : null);

        if (ID) {
            url += "(" + ID + ")";
        } else {
            throw ('[SPHTTP] Delete request data is wrong!');
        }

        return this._rest(url, { method: "POST" });
    }

    public async user(config: any = false): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        if (typeof config === "string") {
            const normalized = config.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
            let url = `_api/web/siteusers?$filter=(startswith(Title,'${config}')) or (startswith(Title,'${normalized}'))`;
            return this._rest(url);
        } else if (typeof config === "number") {
            return this._rest('_api/Web/SiteUserInfoList/Items(' + config + ')');
        } else if (config === true || config === false) {
            return this._rest('_api/Web/CurrentUser?$expand=groups');
        }

        throw ('[SPHTTP] User Request Failed!');
    }

    private async _verifyDigest(refresh: boolean = false): Promise<string | null> {
        const el = document.querySelector("#__REQUESTDIGEST") as HTMLInputElement;
        if (el && el.value && !refresh) {
            this.digest = el.value;
        } else if (refresh || !this.digest) {
            const r = await this._rest("_api/contextinfo", { method: "POST" });
            this.digest = r?.GetContextWebInformation?.FormDigestValue || '';
        }
        return this.digest;
    }

    private _rest(url: string, fetchOptions: any = {}, config: any = {}): Promise<any> {
        const urlFetch = this.baseURL + url;
        return this._fetch(urlFetch, fetchOptions)
            .then((resp: any) => {
                if (resp?.d?.results) {
                    return resp.d.results.map((res: any) => this._cleanObj(res, (config.select || [])));
                } else if (resp?.d) {
                    return this._cleanObj(resp.d, (config.select || []));
                }
                return resp;
            });
    }

    private async _fetch(url: string, fetchOptions: any): Promise<any> {
        const controller = new AbortController();
        const abortFetch = setTimeout(() => controller.abort(), this.timeout);

        try {
            const resp = await fetch(url, Object.assign({
                signal: controller.signal,
                headers: this._headers
            }, fetchOptions));

            clearTimeout(abortFetch);

            let responseContent: any;
            try {
                responseContent = await resp.text();
            } catch (e) {
                return resp.status > 199 && resp.status < 300;
            }

            try {
                const json = JSON.parse(responseContent);
                if (json["odata.error"]?.message?.value) throw json["odata.error"].message.value;
                if (json.error?.message?.value) throw json.error.message.value;
                return fetchOptions?.method === "POST" && json?.d?.Id ? json.d : json;
            } catch (e) {
                return (resp.status > 199 && resp.status < 300) ? (responseContent || true) : false;
            }
        } catch (err) {
            return this._error(err);
        }
    }

    private _cleanObj(obj: any, listKeys: string[] = []): any {
        if (this.cleanResponse && listKeys.length && listKeys[0] !== "") {
            const newObject: any = {};
            for (const key of listKeys) {
                const k = key.trim();
                if (obj[k] !== undefined) {
                    newObject[k] = obj[k];
                }
            }
            return newObject;
        }
        return obj;
    }

    private _recursive(requestUrl: string, obj: any = {}): Promise<any[]> {
        return new Promise((resolve, reject) => {
            const content: any[] = [];
            const loop = async (url: string) => {
                try {
                    const resp = await this._fetch(url, {});
                    if (resp?.d?.results) {
                        content.push(...resp.d.results.map((res: any) => this._cleanObj(res, obj.select.split(","))));
                    }
                    if (resp?.d?.__next) {
                        await loop(resp.d.__next);
                    } else {
                        resolve(content);
                    }
                } catch (e) { reject(e); }
            };
            loop(requestUrl);
        });
    }
}

export default new SpHttp();