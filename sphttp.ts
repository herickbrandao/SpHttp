/** SPHTTP 1.0.3 - TypeScript compatible version */

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

interface IBatchRequest {
    url: string;
    action?: 'GET' | 'POST' | 'UPDATE' | 'DELETE';
    item?: any;
}

interface IHeadersMap {
    [key: string]: string;
}

class SpHttp {
    public baseURL: string = "../";
    public cleanResponse: boolean = true;
    public digest: string | null = null;
    public headers: IHeadersMap = {
        "Accept": "application/json; odata=verbose"
    };
    private _headers: IHeadersMap = {};
    public timeout: number = 15000;
    public top: number = 5000;
    public version: string = "1.0.3";

    private _createUUID(): string {
        // Compatible replacement for crypto.randomUUID()
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c: string) => {
            const r = (Math.random() * 16) | 0;
            const v = c === "x" ? r : ((r & 0x3) | 0x8);
            return v.toString(16);
        });
    }

    private _error(e: any, obj: any = {}): any {
        console.error(e, obj);
        return e;
    }

    private _normalizeSelectExpand(value: string | string[] | undefined): string[] {
        if (!value) return [];
        if (Array.isArray(value)) return value;
        return [value];
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
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        const selectArr = this._normalizeSelectExpand(obj.select);
        const expandArr = this._normalizeSelectExpand(obj.expand);
        const selectStr = selectArr.join();
        const expandStr = expandArr.join();

        if (ID) {
            requestUrl += "(" + ID + ")";

            if (obj.versions) {
                requestUrl += "/versions";
            }

            if (selectStr.length) {
                requestUrl += concatenation + "$select=" + selectStr;
                concatenation = "&";
            }

            if (expandStr.length) {
                requestUrl += concatenation + "$expand=" + expandStr;
                concatenation = "&";
            }

            if (obj.avoidcache) {
                requestUrl += concatenation + "$v=" + this._createUUID();
            }

            return this._rest(requestUrl, {}, {
                select: selectArr
            });
        }

        requestUrl += concatenation + "$top=" + top;
        concatenation = "&";

        if (selectStr.length) {
            requestUrl += concatenation + "$select=" + selectStr;
        }

        if (expandStr.length) {
            requestUrl += concatenation + "$expand=" + expandStr;
        }

        if (obj.filter && obj.filter.length) {
            requestUrl += concatenation + "$filter=" + obj.filter;
        }

        if (obj.avoidcache) {
            requestUrl += concatenation + "$v=" + this._createUUID();
        }

        if (!obj.recursive) {
            if (obj.orderby && obj.orderby.length) {
                requestUrl += concatenation + "$orderby=" + obj.orderby;
            }

            return this._rest(requestUrl, {}, {
                select: selectArr
            });
        }

        return this._recursive(this.baseURL + requestUrl, {
            ...obj,
            select: selectStr
        });
    }

    public async add(list: string, obj: any): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest || ""
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
            "X-RequestDigest": this.digest || "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = obj && (obj.ID || obj.Id || null);

        if (typeof obj === "object" && ID) {
            url += "(" + ID + ")";
        } else {
            throw ("[ERROR] Put request data is wrong!");
        }

        return this._rest(url, {
            method: "POST",
            body: typeof obj === "object" ? JSON.stringify(obj) : obj
        });
    }

    public async recycle(list: string, obj: any): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = (typeof obj === "object") ? (obj.ID || obj.Id || null) : obj;

        if (typeof obj === "object" && ID) {
            url += "(" + ID + ")/recycle()";
        } else if (typeof obj === "number") {
            url += "(" + obj + ")/recycle()";
        } else {
            throw ("[SPHTTP] Recycle request data is wrong!");
        }

        return this._rest(url, {
            method: "POST"
        });
    }

    public async delete(list: string, obj: any): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest || "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        }, this.headers);

        let url = "_api/lists/getbytitle('" + list + "')/items";
        const ID = (typeof obj === "object") ? (obj.ID || obj.Id || null) : obj;

        if (typeof obj === "object" && ID) {
            url += "(" + ID + ")";
        } else if (typeof obj === "number") {
            url += "(" + obj + ")";
        } else {
            throw ("[SPHTTP] Delete request data is wrong!");
        }

        return this._rest(url, {
            method: "POST"
        });
    }

    public async attach(list: string, config: ISpHttpConfig = {}): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        const ID = config.ID || config.Id || (typeof config === "number" ? config : null) || null;

        if (!window.FileReader) {
            throw ("[SPHTTP] Your browser does not have support for FileReader!");
        }

        if (!ID) {
            throw ("[SPHTTP] List ID not found at attach request!");
        }

        if (!config.target && !config.delete) {
            return this._rest("_api/lists/getbytitle('" + list + "')/items(" + ID + ")/AttachmentFiles");
        }

        if (config.delete) {
            this._headers = Object.assign({
                "X-HTTP-Method": "DELETE"
            }, this._headers);

            const delUrl = "_api/lists/GetByTitle('" + list + "')/items(" + ID + ")/AttachmentFiles/getByFileName('" + config.delete + "')";
            return this._rest(delUrl, {
                method: "DELETE"
            });
        }

        const input = typeof config.target === "string" ? (document.querySelector(config.target) as HTMLInputElement | null) : null;
        const files = input && input.files ? input.files : null;

        if (!files || files.length === 0) {
            return this._rest("_api/lists/getbytitle('" + list + "')/items(" + ID + ")/AttachmentFiles");
        }

        const appends: Promise<ArrayBuffer>[] = [];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (file) {
                appends.push(this._getFileBuffer(file));
            }
        }

        const buffers = await Promise.all(appends);
        const results: any[] = [];

        for (let i = 0; i < buffers.length; i++) {
            const file = files[i];
            if (!file) continue;

            this._headers = Object.assign({
                "X-RequestDigest": this.digest || "",
                "content-length": String(buffers[i].byteLength)
            }, this.headers);

            const url = "_api/lists/GetByTitle('" + list + "')/items(" + ID + ")/AttachmentFiles/add(FileName='" + file.name + "')";
            const res = await this._rest(url, {
                method: "POST",
                body: buffers[i]
            });

            results.push(res && res.d ? res.d : res);
        }

        return results;
    }

    public async attachDoc(config: ISpHttpConfig = {}): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        if (!window.FileReader) {
            throw ("[SPHTTP] Your browser does not have support for FileReader!");
        }

        if (config.delete) {
            this._headers = Object.assign({
                "X-HTTP-Method": "DELETE"
            }, this._headers);

            let url = "_api/web/getfilebyserverrelativeurl('" + config.library + "/" + config.delete + "')";
            if (config.recycle && config.recycle === true) {
                url += "/recycle()";
            }

            return this._rest(url, {
                method: "POST"
            });
        }

        const input = typeof config.target === "string" ? (document.querySelector(config.target) as HTMLInputElement | null) : null;
        const files = input && input.files ? input.files : null;

        if (!files || files.length === 0) {
            let tname = "";
            if (config.startswith && config.name) {
                tname = "?$filter=startswith(Name,'" + config.name + "')";
            }

            return this._rest("_api/web/GetFolderByServerRelativeUrl('" + config.library + "')/Files" + tname);
        }

        const appends: Promise<ArrayBuffer>[] = [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (file) {
                appends.push(this._getFileBuffer(file));
            }
        }

        const buffers = await Promise.all(appends);
        const results: any[] = [];

        for (let i = 0; i < buffers.length; i++) {
            const file = files[i];
            if (!file) continue;

            this._headers = Object.assign({
                "content-length": String(buffers[i].byteLength)
            }, this._headers);

            const name = config.name ? config.name : file.name;
            const url = "_api/web/GetFolderByServerRelativeUrl('" + config.library + "')/Files/Add(url='" + name + "', overwrite=true)";

            const res = await this._rest(url, {
                method: "POST",
                body: buffers[i]
            });

            results.push(res && res.d ? res.d : res);
        }

        return results;
    }

    public async user(config: any = false): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        switch (typeof config) {
            case "string": {
                const searches = [
                    config.toLowerCase(),
                    config.toUpperCase()
                ];

                // Capitalize
                const arr = searches[0].split(" ");
                for (let i = 0; i < arr.length; i++) {
                    arr[i] = arr[i].charAt(0).toUpperCase() + arr[i].slice(1);
                }
                searches[2] = arr.join(" ");

                // Normalize
                searches[3] = config.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

                let url = "_api/web/siteusers?$filter=(";
                url += "(startswith(Title,'" + searches[0] + "')) or ";
                url += "(startswith(Title,'" + searches[1] + "')) or ";
                url += "(startswith(Title,'" + searches[3] + "')) or ";
                url += "(startswith(Title,'" + searches[2] + "')))";

                return this._rest(url);
            }

            case "number":
                return this._rest("_api/Web/SiteUserInfoList/Items(" + config + ")");

            case "object":
                if (config && config.ID) {
                    return this._rest("_api/web/GetUserById(" + config.ID + ")/Groups");
                }
                break;

            case "boolean":
                return this._rest("_api/Web/CurrentUser?$expand=groups");
        }

        throw ("[SPHTTP] User Request Failed!");
    }

    public async batch(batchRequests: IBatchRequest[]): Promise<any> {
        await this._verifyDigest();

        const batchBoundary = this._createUUID();
        const changesetBoundary = this._createUUID();
        const propData: string[] = [];
        const changesetRequests: IBatchRequest[] = [];
        const getRequests: IBatchRequest[] = [];

        for (let i = 0; i < batchRequests.length; i++) {
            const req = batchRequests[i];
            req.action = (req.action || "GET").toUpperCase() as any;

            if (req.action === "GET") {
                getRequests.push(req);
            } else {
                changesetRequests.push(req);
            }
        }

        if (changesetRequests.length) {
            propData.push("--batch_" + batchBoundary);
            propData.push('Content-Type: multipart/mixed; boundary="changeset_' + changesetBoundary + '"');
            propData.push("Content-Transfer-Encoding: binary");
            propData.push("");

            for (let i = 0; i < changesetRequests.length; i++) {
                const req = changesetRequests[i];

                propData.push("--changeset_" + changesetBoundary);
                propData.push("Content-Type: application/http");
                propData.push("Content-Transfer-Encoding: binary");
                propData.push("");

                switch (req.action) {
                    case "UPDATE":
                        propData.push("PATCH " + req.url + " HTTP/1.1");
                        propData.push("If-Match: *");
                        propData.push("Content-Type: application/json;odata=verbose");
                        propData.push("");
                        propData.push(JSON.stringify(req.item || {}));
                        propData.push("");
                        break;

                    case "DELETE":
                        propData.push("DELETE " + req.url + " HTTP/1.1");
                        propData.push("If-Match: *");
                        propData.push("");
                        break;

                    case "POST":
                        propData.push("POST " + req.url + " HTTP/1.1");
                        propData.push("Content-Type: application/json;odata=verbose");
                        propData.push("");

                        if (req.item) {
                            propData.push(JSON.stringify(req.item || {}));
                            propData.push("");
                        }
                        break;
                }
            }

            propData.push("--changeset_" + changesetBoundary + "--");
        }

        for (let i = 0; i < getRequests.length; i++) {
            const req = getRequests[i];

            propData.push("--batch_" + batchBoundary);
            propData.push("Content-Type: application/http");
            propData.push("Content-Transfer-Encoding: binary");
            propData.push("");
            propData.push("GET " + req.url + " HTTP/1.1");
            propData.push("Accept: application/json;odata=verbose");
            propData.push("");
        }

        propData.push("--batch_" + batchBoundary + "--");

        const bodyBoundary = propData.join("\r\n");
        const requestHeaders = {
            "X-RequestDigest": this.digest || "",
            "Content-Type": 'multipart/mixed; boundary="batch_' + batchBoundary + '"'
        };

        const response = await fetch(this.baseURL + "_api/$batch", {
            method: "POST",
            headers: requestHeaders,
            credentials: "include",
            body: bodyBoundary
        });

        const raw = await response.text();

        if (!raw || typeof raw !== "string") {
            return { ok: false, data: raw };
        }

        const responses: any[] = [];
        const parsedRaw = raw.split("\r\n\r\n");

        for (let i = 0; i < parsedRaw.length; i++) {
            let p = parsedRaw[i];

            if (p.indexOf("No Content") > -1) {
                if (p.indexOf("HTTP") === 0 && p.indexOf("204") > -1) {
                    responses.push({ ok: true, data: true });
                } else {
                    responses.push({ ok: false, data: false });
                }
            } else if (p.indexOf("<") === 0) {
                let parsedCtx: any = {};

                try {
                    if (p.indexOf("<m:error") > -1) {
                        p = p.split("<m:message")[1].split(">")[1].split("</m:message")[0];
                        parsedCtx = {
                            ok: false,
                            data: p
                        };
                    } else {
                        parsedCtx = {
                            ok: true,
                            data: p.split(">\r\n")[0] + ">"
                        };
                    }
                } catch (e) {
                    parsedCtx = {
                        ok: false,
                        data: p
                    };
                }

                responses.push(parsedCtx);
            } else if (p.indexOf("{") === 0) {
                try {
                    const parsedJson = JSON.parse(p.split("}\r\n")[0] + "}");
                    responses.push({
                        ok: true,
                        data: parsedJson
                    });
                } catch (e) {
                    responses.push({
                        ok: false,
                        data: p
                    });
                }
            }
        }

        return responses;
    }

    public async rest(url: string, fetchOptions: RequestInit = {}, config: any = {}): Promise<any> {
        await this._verifyDigest();

        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest || ""
        }, this.headers);

        return this._rest(url, fetchOptions, config);
    }

    public async _verifyDigest(refresh: boolean = false): Promise<string> {
        const digestEl = document.querySelector("#__REQUESTDIGEST") as HTMLInputElement | null;

        if (digestEl && digestEl.value && !refresh) {
            this.digest = digestEl.value;
        } else if (refresh || !this.digest) {
            try {
                const r = await this._rest("_api/contextinfo", {
                    method: "POST"
                });

                this.digest = (r && r.GetContextWebInformation && r.GetContextWebInformation.FormDigestValue) || "";
                return this.digest || "";
            } catch (e) {
                this._error(e);
                return "";
            }
        }

        return this.digest || "";
    }

    public _rest(url: string, fetchOptions: RequestInit = {}, config: any = {}): Promise<any> {
        const fullUrl = this.baseURL + url;

        return this._fetch(fullUrl, fetchOptions).then((resp: any) => {
            if (resp && resp.d && resp.d.results) {
                return resp.d.results.map((res: any) => {
                    return this._cleanObj(res, config.select || []);
                });
            } else if (resp && resp.d) {
                return this._cleanObj(resp.d, config.select || []);
            }

            return resp;
        });
    }

    public async _fetch(url: string, fetchOptions: RequestInit = {}): Promise<any> {
        const controller = new AbortController();
        const abortFetch = setTimeout(() => {
            controller.abort();
        }, this.timeout);

        const requestOptions: RequestInit = Object.assign({
            signal: controller.signal,
            headers: this._headers
        }, fetchOptions);

        try {
            const resp = await fetch(url, requestOptions);
            let responseContent: any = false;

            try {
                responseContent = await resp.text();
            } catch (e) {
                clearTimeout(abortFetch);
                if (resp.status > 199 && resp.status < 300) {
                    return true;
                }
                return false;
            }

            let parsed: any = responseContent;
            try {
                parsed = JSON.parse(responseContent);
            } catch (e) {
                if (resp.status > 199 && resp.status < 300) {
                    clearTimeout(abortFetch);
                    return responseContent && responseContent.length ? responseContent : true;
                }

                this._error("[SPHTTP] Internal error - The response was not a JSON.", [resp, fetchOptions]);
                clearTimeout(abortFetch);
                return false;
            }

            if (parsed && parsed["odata.error"] && parsed["odata.error"].message && parsed["odata.error"].message.value) {
                clearTimeout(abortFetch);
                throw new Error(parsed["odata.error"].message.value);
            }

            if (parsed && parsed.error && parsed.error.message && parsed.error.message.value) {
                clearTimeout(abortFetch);
                throw new Error(parsed.error.message.value);
            }

            if (fetchOptions && (fetchOptions.method === "POST") && parsed && parsed.d && parsed.d.Id) {
                clearTimeout(abortFetch);
                return parsed.d;
            }

            clearTimeout(abortFetch);
            return parsed;
        } catch (e) {
            clearTimeout(abortFetch);
            return this._error(e);
        }
    }

    public _cleanObj(obj: any, listKeys: string[] = []): any {
        if (this.cleanResponse && listKeys.length && listKeys[0] !== "") {
            const newObject: any = {};

            if (listKeys.join().length > 0) {
                for (let i = 0; i < listKeys.length; i++) {
                    const key = listKeys[i];

                    if (obj[key]) {
                        newObject[key] = obj[key];
                    } else {
                        const parts = key.split("/");
                        const rootKey = parts[0];
                        const childKey = parts[1];

                        if (obj[rootKey]) {
                            if (!newObject[rootKey]) {
                                newObject[rootKey] = [];
                            }

                            if (obj[rootKey].results && !newObject[rootKey].length) {
                                newObject[rootKey].push.apply(newObject[rootKey], obj[rootKey].results);
                            } else if (obj[rootKey] && childKey && obj[rootKey][childKey]) {
                                newObject[rootKey][0] = obj[rootKey];
                            }
                        } else {
                            newObject[key] = typeof newObject[key] === "object" ? null : newObject[key];
                        }
                    }
                }
            }

            return newObject;
        }

        return obj;
    }

    public _reset(): void {
        this.baseURL = "../";
        this.cleanResponse = true;
        this.digest = null;
        this.headers = {
            "Accept": "application/json; odata=verbose"
        };
        this._headers = {};
        this.timeout = 15000;
        this.top = 5000;
    }

    public _getFileBuffer(file: File): Promise<ArrayBuffer> {
        return new Promise((resolve, reject) => {
            try {
                const reader = new FileReader();

                reader.onload = (e: any) => {
                    resolve(e.target.result as ArrayBuffer);
                };

                reader.onerror = (e: any) => {
                    reject(e.target.error);
                };

                reader.readAsArrayBuffer(file);
            } catch (error) {
                console.error("[ERROR] " + error);
                reject(error);
            }
        });
    }

    public _recursive(requestUrl: string, obj: ISpHttpConfig = {}): Promise<any> {
        const _this = this;
        return new Promise((resolve, reject) => {
            const content: any[] = [];

            const loop = async function (url: string): Promise<void> {
                _this._fetch(url, {})
                    .then(async function (resp: any) {
                        const selectArr = _this._normalizeSelectExpand(obj.select);

                        if (resp && resp.d && resp.d.results) {
                            content.push.apply(
                                content,
                                resp.d.results.map(function (res: any) {
                                    return _this._cleanObj(res, selectArr);
                                })
                            );
                        } else if (resp && resp.d) {
                            content.push.apply(content, _this._cleanObj(resp.d, selectArr));
                        }

                        if (resp && resp.d && resp.d.__next) {
                            await loop(resp.d.__next);
                        } else {
                            resolve(content);
                        }
                    })
                    .catch(function (error: any) {
                        reject(error);
                    });
            };

            loop(requestUrl);
        });
    }
}

export default new SpHttp();
