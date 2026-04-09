/** SPHTTP 1.0.3 - Versão Completa TypeScript */

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
    delete?: string;
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

    // --- MÉTODOS DE ITENS ---

    public async items(list: string, obj: ISpHttpConfig = {}): Promise<any> {
        if (!list) return this._error("Your object has not a list name", obj);

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

        const selectArr = Array.isArray(obj.select) ? obj.select : (obj.select ? [obj.select] : []);
        const expandArr = Array.isArray(obj.expand) ? obj.expand : (obj.expand ? [obj.expand] : []);
        const selectStr = selectArr.join();
        const expandStr = expandArr.join();

        if (ID) {
            requestUrl += "(" + ID + ")";
            if (obj.versions) requestUrl += "/versions";
            if (selectStr.length) { requestUrl += concatenation + "$select=" + selectStr; concatenation = "&"; }
            if (expandStr.length) { requestUrl += concatenation + "$expand=" + expandStr; concatenation = "&"; }
            if (obj.avoidcache) requestUrl += concatenation + "$v=" + (window as any).crypto.randomUUID();
            return this._rest(requestUrl, {}, { select: selectArr });
        }

        requestUrl += concatenation + "$top=" + top;
        concatenation = "&";

        if (selectStr.length) requestUrl += concatenation + "$select=" + selectStr;
        if (expandStr.length) requestUrl += concatenation + "$expand=" + expandStr;
        if (obj.filter?.length) requestUrl += concatenation + "$filter=" + obj.filter;
        if (obj.avoidcache) requestUrl += concatenation + "$v=" + (window as any).crypto.randomUUID();

        if (!obj.recursive) {
            if (obj.orderby?.length) requestUrl += concatenation + "$orderby=" + obj.orderby;
            return this._rest(requestUrl, {}, { select: selectArr });
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
        if (!ID) throw ('[ERROR] Update ID is missing!');

        url += "(" + ID + ")";
        return this._rest(url, {
            method: "POST",
            body: typeof obj === 'object' ? JSON.stringify(obj) : obj
        });
    }

    public async recycle(list: string, obj: any): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = typeof obj === 'object' ? (obj.ID || obj.Id) : obj;
        if (!ID) throw ("[SPHTTP] Recycle ID is missing!");

        url += "(" + ID + ")/recycle()";
        return this._rest(url, { method: "POST" });
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
        const ID = typeof obj === 'object' ? (obj.ID || obj.Id) : obj;
        if (!ID) throw ('[SPHTTP] Delete ID is missing!');

        url += "(" + ID + ")";
        return this._rest(url, { method: "POST" });
    }

    // --- ARQUIVOS E ANEXOS ---

    public async attach(list: string, config: ISpHttpConfig = {}): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({ "X-RequestDigest": this.digest }, this.headers);

        const ID = config.ID || config.Id || (typeof config === 'number' ? config : null);
        if (!window.FileReader) throw ('[SPHTTP] FileReader not supported!');
        if (!ID) throw ('[SPHTTP] List ID not found!');

        if (config.delete) {
            this._headers["X-HTTP-Method"] = "DELETE";
            const url = `_api/lists/GetByTitle('${list}')/items(${ID})/AttachmentFiles/getByFileName('${config.delete}')`;
            return this._rest(url, { method: "DELETE" });
        }

        const input = config.target ? document.querySelector(config.target) as HTMLInputElement : null;
        const files = input?.files;
        if (!files) return this._rest(`_api/lists/getbytitle('${list}')/items(${ID})/AttachmentFiles`);

        const appends: Promise<any>[] = [];
        for (let i = 0; i < files.length; i++) {
            appends.push(this._getFileBuffer(files[i]));
        }

        const buffers = await Promise.all(appends);
        const results = [];
        for (let i = 0; i < buffers.length; i++) {
            this._headers["content-length"] = buffers[i].byteLength;
            const url = `_api/lists/GetByTitle('${list}')/items(${ID})/AttachmentFiles/add(FileName='${files[i].name}')`;
            const res = await this._rest(url, { method: "POST", body: buffers[i] });
            results.push(res?.d || res);
        }
        return results;
    }

    public async attachDoc(config: ISpHttpConfig = {}): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({ "X-RequestDigest": this.digest }, this.headers);

        if (config.delete) {
            this._headers["X-HTTP-Method"] = "DELETE";
            let url = `_api/web/getfilebyserverrelativeurl('${config.library}/${config.delete}')`;
            if (config.recycle) url += '/recycle()';
            return this._rest(url, { method: "POST" });
        }

        const input = config.target ? document.querySelector(config.target) as HTMLInputElement : null;
        const files = input?.files;
        if (!files) {
            let tname = config.startswith ? `?$filter=startswith(Name,'${config.name}')` : '';
            return this._rest(`_api/web/GetFolderByServerRelativeUrl('${config.library}')/Files${tname}`);
        }

        const appends = [];
        for (let i = 0; i < files.length; i++) appends.push(this._getFileBuffer(files[i]));
        
        const buffers = await Promise.all(appends);
        const results = [];
        for (let i = 0; i < buffers.length; i++) {
            this._headers["content-length"] = buffers[i].byteLength;
            const name = config.name || files[i].name;
            const url = `_api/web/GetFolderByServerRelativeUrl('${config.library}')/Files/Add(url='${name}', overwrite=true)`;
            const res = await this._rest(url, { method: "POST", body: buffers[i] });
            results.push(res?.d || res);
        }
        return results;
    }

    // --- USUÁRIOS E BATCH ---

    public async user(config: any = false): Promise<any> {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        if (typeof config === "string") {
            const searches = [config.toLowerCase(), config.toUpperCase()];
            const capitalized = config.split(" ").map((s:any) => s.charAt(0).toUpperCase() + s.slice(1)).join(" ");
            const normalized = config.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
            
            let filter = `(startswith(Title,'${searches[0]}')) or (startswith(Title,'${searches[1]}')) or (startswith(Title,'${normalized}')) or (startswith(Title,'${capitalized}'))`;
            return this._rest(`_api/web/siteusers?$filter=${filter}`);
        }
        if (typeof config === "number") return this._rest(`_api/Web/SiteUserInfoList/Items(${config})`);
        if (typeof config === "object" && config.ID) return this._rest(`_api/web/GetUserById(${config.ID})/Groups`);
        if (typeof config === "boolean") return this._rest('_api/Web/CurrentUser?$expand=groups');

        throw ('[SPHTTP] User Request Failed!');
    }

    public async batch(batchRequests: any[]): Promise<any> {
        await this._verifyDigest();
        const batchBoundary = (window as any).crypto.randomUUID();
        const changesetBoundary = (window as any).crypto.randomUUID();
        const propData: string[] = [];
        
        const changeset = batchRequests.filter(r => r.action && r.action.toUpperCase() !== "GET");
        const gets = batchRequests.filter(r => !r.action || r.action.toUpperCase() === "GET");

        if (changeset.length) {
            propData.push(`--batch_${batchBoundary}`, `Content-Type: multipart/mixed; boundary="changeset_${changesetBoundary}"`, `Content-Transfer-Encoding: binary`, '');
            for (const r of changeset) {
                propData.push(`--changeset_${changesetBoundary}`, `Content-Type: application/http`, `Content-Transfer-Encoding: binary`, '');
                const action = r.action.toUpperCase();
                if (action === "UPDATE") {
                    propData.push(`PATCH ${r.url} HTTP/1.1`, `If-Match: *`, `Content-Type: application/json;odata=verbose`, '', JSON.stringify(r.item || {}), '');
                } else if (action === "DELETE") {
                    propData.push(`DELETE ${r.url} HTTP/1.1`, `If-Match: *`, '');
                } else if (action === "POST") {
                    propData.push(`POST ${r.url} HTTP/1.1`, `Content-Type: application/json;odata=verbose`, '', JSON.stringify(r.item || {}), '');
                }
            }
            propData.push(`--changeset_${changesetBoundary}--`);
        }

        for (const r of gets) {
            propData.push(`--batch_${batchBoundary}`, `Content-Type: application/http`, `Content-Transfer-Encoding: binary`, '', `GET ${r.url} HTTP/1.1`, `Accept: application/json;odata=verbose`, '');
        }
        propData.push(`--batch_${batchBoundary}--`);

        const res = await fetch(this.baseURL + "_api/$batch", {
            method: "POST",
            headers: { 'X-RequestDigest': this.digest!, 'Content-Type': `multipart/mixed; boundary="batch_${batchBoundary}"` },
            body: propData.join('\r\n')
        });

        const raw = await res.text();
        return raw.split("--batch").filter(p => p.includes("{") || p.includes("<")).map(p => {
            try { return { ok: true, data: JSON.parse(p.substring(p.indexOf("{"), p.lastIndexOf("}") + 1)) }; }
            catch { return { ok: false, data: p }; }
        });
    }

    // --- INTERNOS ---

    public async _verifyDigest(refresh: boolean = false): Promise<string> {
        const el = document.querySelector("#__REQUESTDIGEST") as HTMLInputElement;
        if (el?.value && !refresh) {
            this.digest = el.value;
        } else if (refresh || !this.digest) {
            const r = await this._rest("_api/contextinfo", { method: "POST" });
            this.digest = r?.GetContextWebInformation?.FormDigestValue || '';
        }
        return this.digest || '';
    }

    public _rest(url: string, fetchOptions: any = {}, config: any = {}): Promise<any> {
        return this._fetch(this.baseURL + url, fetchOptions).then((resp: any) => {
            if (resp?.d?.results) return resp.d.results.map((res: any) => this._cleanObj(res, config.select || []));
            if (resp?.d) return this._cleanObj(resp.d, config.select || []);
            return resp;
        });
    }

    public async _fetch(url: string, fetchOptions: any): Promise<any> {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), this.timeout);
        try {
            const resp = await fetch(url, { ...fetchOptions, signal: controller.signal, headers: this._headers });
            clearTimeout(timeout);
            const text = await resp.text();
            try {
                const json = JSON.parse(text);
                if (json["odata.error"] || json.error) throw (json["odata.error"] || json.error).message.value;
                return json;
            } catch { return resp.ok ? (text || true) : false; }
        } catch (e) { return this._error(e); }
    }

    public _cleanObj(obj: any, keys: string[]): any {
        if (!this.cleanResponse || !keys.length || keys[0] === "") return obj;
        const newObj: any = {};
        for (const key of keys) {
            const k = key.trim();
            if (obj[k] !== undefined) newObj[k] = obj[k];
        }
        return newObj;
    }

    public _getFileBuffer(file: File): Promise<ArrayBuffer> {
        return new Promise((res, rej) => {
            const reader = new FileReader();
            reader.onload = (e: any) => res(e.target.result);
            reader.onerror = (e: any) => rej(e.target.error);
            reader.readAsArrayBuffer(file);
        });
    }

    public _recursive(requestUrl: string, obj: any): Promise<any> {
        return new Promise((res, rej) => {
            const content: any[] = [];
            const loop = async (url: string) => {
                try {
                    const resp = await this._fetch(url, {});
                    if (resp?.d?.results) content.push(...resp.d.results.map((r: any) => this._cleanObj(r, obj.select.split(","))));
                    if (resp?.d?.__next) await loop(resp.d.__next);
                    else res(content);
                } catch (e) { rej(e); }
            };
            loop(requestUrl);
        });
    }
}

export default new SpHttp();
