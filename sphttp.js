/** SPHTTP 1.0.2 - https://github.com/herickbrandao/SpHttp */
var sphttp = {
    baseURL: "../",
    cleanResponse: true,
    digest: null,
    headers: {
        "Accept": "application/json; odata=verbose"
    },
    _headers: {},
    timeout: 15000,
    top: 5000,
    version: "1.0.2",

    async items(list, obj = {}) {
        if(!list) {
            return this._error("Your object has not a list name", obj);
        }

        const ID = obj.ID || obj.Id || false;
        const top = obj.top || this.top || 5000;
        var requestUrl = "_api/lists/getbytitle('" + list + "')/items";
        var concatenation = "?";

        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        if(!obj.select) {
            obj.select = [];
        }
        if(!obj.expand) {
            obj.expand = [];
        }
        if(typeof obj.select === "object") {
            obj.select = obj.select.join();
        }
        if(typeof obj.expand === "object") {
            obj.expand = obj.expand.join();
        }

        if(ID) {
            requestUrl += "(" + ID + ")";
            if(obj.versions) {
                requestUrl += "/versions";
            }
            if(obj.select?.length) {
                requestUrl += concatenation + "$select=" + obj.select;
                concatenation = "&";
            }
            if(obj.expand?.length > 0) {
                requestUrl += concatenation + "$expand=" + obj.expand;
                concatenation = "&";
            }
            if(obj.avoidcache) {
                requestUrl += concatenation + "$v=" + window.crypto.randomUUID();
                concatenation = "&";
            }
            return this._rest(requestUrl, {}, {
                select: obj.select.split(",")
            });
        }

        requestUrl += concatenation + "$top=" + top;
        concatenation = "&";

        if(obj.select?.length) {
            requestUrl += concatenation + "$select=" + obj.select;
        }
        if(obj.expand?.length) {
            requestUrl += concatenation + "$expand=" + obj.expand;
        }
        if(obj.filter?.length) {
            requestUrl += concatenation + "$filter=" + obj.filter;
        }
        if(obj.avoidcache) {
            requestUrl += concatenation + "$v=" + window.crypto.randomUUID();
        }

        if(!obj.recursive) {
            if(obj.orderby?.length) {
                requestUrl += concatenation + "$orderby=" + obj.orderby;
            }
            return this._rest(requestUrl, {}, {
                select: obj.select.split(",")
            });
        }

        return this._recursive(this.baseURL + requestUrl, obj);
    },

    async add(list, obj) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest
        }, this.headers);

        var url = "_api/lists/getbytitle('" + list + "')/items";
        return this._rest(url, {
            method: "POST",
            body: typeof obj === "object" ? JSON.stringify(obj) : obj
        });
    },

    async update(list, obj) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        }, this.headers);

        var url = "_api/lists/getbytitle('" + list + "')/items";
        var ID = obj.ID || obj.Id || null;

        if(typeof obj === 'object' && ID) {
            url += "(" + ID + ")";
        } else {
            throw ('[ERROR] Put request data is wrong!');
        }

        return this._rest(url, {
            method: "POST",
            body: typeof obj === 'object' ? JSON.stringify(obj) : obj
        });
    },

    async recycle(list, obj) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
        }, this.headers);

        var url = "_api/lists/getbytitle('" + list + "')/items";
        var ID = obj.ID || obj.Id || null;

        if(typeof obj === 'object' && ID) {
            url += "(" + ID + ")/recycle()";
        } else if(typeof obj === 'number') {
            url += "(" + obj + ")/recycle()";
        } else {
            throw ("[SPHTTP] Recycle request data is wrong!");
        }

        return this._rest(url, {
            method: "POST",
            body: typeof obj === 'object' ? JSON.stringify(obj) : obj
        });
    },

    async delete(list, obj) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": this.digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
        }, this.headers);

        var url = "_api/lists/getbytitle('" + list + "')/items";
        var ID = obj.ID || obj.Id || null;

        if(typeof obj === 'object' && ID) {
            url += "(" + ID + ")";
        } else if(typeof obj === 'number') {
            url += "(" + obj + ")";
        } else {
            throw ('[SPHTTP] Delete request data is wrong!');
        }

        return this._rest(url, {
            method: "POST"
        });
    },

    async attach(list, config = {}) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "X-RequestDigest": this.digest,
        }, this.headers);

        var ID = config.ID || config.Id || (typeof config === 'number' ? config : null) || null;
        if(!window.FileReader) {
            throw ('[SPHTTP] Your browser does not have support for FileReader!');
        }
        if(isNaN(parseInt(ID))) {
            throw ('[SPHTTP] List ID not found at attach request!');
        }
        if(!config.target && !config.delete) {
            return this._rest("_api/lists/getbytitle('" + list + "')/items(" + ID + ")/AttachmentFiles");
        }

        if(config.delete) {
            this._headers = Object.assign({
                "X-HTTP-Method": "DELETE",
            }, this._headers);

            url = "_api/lists/GetByTitle('" + list + "')/items(" + ID + ")/AttachmentFiles/getByFileName('" + config.delete + "')";
            return this._rest(url, {
                method: "DELETE"
            });
        }

        var items = typeof config.target === "string" && document.querySelector(config.target) ? document.querySelector(config.target).files : false;
        var appends = [];

        for(var i in items) {
            if(typeof items[i] === 'object') {
                appends.push(this._getFileBuffer(items[i]));
            }
        }

        return Promise.all(appends).then(function(promises) {
            var httpRes = [],
                url = '';

            for(var i in promises) {
                this._headers = Object.assign({
                    "X-RequestDigest": this.digest,
                    "content-length": promises[i].byteLength
                }, this.headers);

                var bytes = new Uint8Array(promises[i]);
                var name = items[i].name;
                url = "_api/lists/GetByTitle('" + listName + "')/items(" + ID + ")/AttachmentFiles/add(FileName='" + name + "')";

                httpRes.push(this._fetch(url, {
                    method: "POST",
                    body: promises[i]
                }));
            }

            return Promise.all(httpRes).then(function(res) {
                for(var i in res) {
                    if(res[i] && res[i].d) {
                        res[i] = res[i].d;
                    }
                }
                return res;
            });
        });
    },

    async attachDoc(config = {}) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "X-RequestDigest": this.digest,
        }, this.headers);

        if(!window.FileReader) {
            throw ('[SPHTTP] Your browser does not have support for FileReader!');
        }
        if(!config.target && !config.delete) {
            if(config.startswith && config.name) {
                config.tname = "?$filter=startswith(Name,'" + config.name + "')";
            }
            return this._rest("_api/web/GetFolderByServerRelativeUrl('" + config.library + "')/Files" + config.tname);
        }

        if(config.delete) {
            this._headers = Object.assign({
                "X-HTTP-Method": "DELETE",
            }, this._headers);

            url = "_api/web/getfilebyserverrelativeurl('" + config.library + '/' + config.delete + "')";
            if(config.recycle && config.recycle == true) {
                url += '/recycle()';
            }
            return this._rest(url, {
                method: "POST"
            });
        }

        var items = typeof config.target === "string" && document.querySelector(config.target) ? document.querySelector(config.target).files : false;
        var appends = [];

        for(var i in items) {
            if(typeof items[i] === 'object') {
                appends.push(this._getFileBuffer(items[i]));
            }
        }

        return Promise.all(appends).then(function(promises) {
            var httpRes = [],
                url = '';

            for(var i in promises) {
                this._headers = Object.assign({
                    "content-length": promises[i].byteLength
                }, this._headers);

                var bytes = new Uint8Array(promises[i]);
                var name = config.name ? config.name : items[i].name;
                url = "_api/web/GetFolderByServerRelativeUrl('" + config.library + "')/Files/Add(url='" + name + "', overwrite=true)";

                httpRes.push(this._rest(url, {
                    method: "POST",
                    body: promises[i]
                }));
            }

            return Promise.all(httpRes).then(function(res) {
                for(var i in res) {
                    if(res[i] && res[i].d) {
                        res[i] = res[i].d;
                    }
                }
                return res;
            });
        });
    },

    async user(config = false) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        switch(typeof config) {
            case "string":
                var searches = [
                    config.toLowerCase(),
                    config.toUpperCase(),
                ];

                // Capitalize
                const str = searches[0];
                const arr = str.split(" ");
                for(var i = 0; i < arr.length; i++) {
                    arr[i] = arr[i].charAt(0).toUpperCase() + arr[i].slice(1);
                }
                searches[2] = arr.join(" ");

                // Normalize
                searches[3] = config.normalize('NFD').replace(/[\u0300-\u036f]/g, "");

                var url = '_api/web/siteusers?$filter=(';
                url += "(startswith(Title,'" + searches[0] + "')) or ";
                url += "(startswith(Title,'" + searches[1] + "')) or ";
                url += "(startswith(Title,'" + searches[3] + "')) or ";
                url += "(startswith(Title,'" + searches[2] + "')))";
                return this._rest(url);
                break;
            case "number":
                return this._rest('_api/Web/SiteUserInfoList/Items(' + config + ')');
                break;
            case "object":
                if(config.ID) {
                    return this._rest("_api/web/GetUserById(" + config.ID + ")/Groups");
                }
                break;
            case "boolean":
                return this._rest('_api/Web/CurrentUser?$expand=groups');
                break;
        }

        throw ('[SPHTTP] User Request Failed!');
    },

    async batch(batchRequests) {
        await this._verifyDigest();
        var batchBoundary = window.crypto.randomUUID();
        var changesetBoundary = window.crypto.randomUUID();
        var propData = [],
            actual = false;
        var changesetRequests = [],
            getRequests = [];

        for(var i in batchRequests) {
            batchRequests[i].action = batchRequests[i].action?.toUpperCase() || "GET";
            if(batchRequests[i].action == "GET") {
                getRequests.push(batchRequests[i]);
            } else {
                changesetRequests.push(batchRequests[i]);
            }
        }

        if(changesetRequests.length) {
            propData.push('--batch_' + batchBoundary);
            propData.push(`Content-Type: multipart/mixed; boundary="changeset_${changesetBoundary}"`);
            propData.push('Content-Transfer-Encoding: binary');
            propData.push('');

            for(var i in changesetRequests) {
                propData.push("--changeset_" + changesetBoundary);
                propData.push('Content-Type: application/http');
                propData.push('Content-Transfer-Encoding: binary');
                propData.push('');

                switch(changesetRequests[i].action) {
                    case "UPDATE":
                        propData.push("PATCH " + changesetRequests[i].url + " HTTP/1.1");
                        propData.push("If-Match: *");
                        propData.push("Content-Type: application/json;odata=verbose");
                        propData.push("");
                        propData.push(JSON.stringify(changesetRequests[i].item || {}));
                        propData.push("");
                        break;
                    case "DELETE":
                        propData.push("DELETE " + changesetRequests[i].url + " HTTP/1.1");
                        propData.push("If-Match: *");
                        propData.push("");
                        break;
                    case "POST":
                        propData.push('POST ' + changesetRequests[i].url + ' HTTP/1.1');
                        propData.push('Content-Type: application/json;odata=verbose');
                        propData.push('');

                        if(changesetRequests[i].item) {
                            propData.push(JSON.stringify(changesetRequests[i].item || {}));
                            propData.push('');
                        }
                        break;
                }
            }

            propData.push('--changeset_' + changesetBoundary + '--');
        }

        for(var i in getRequests) {
            propData.push('--batch_' + batchBoundary);
            propData.push('Content-Type: application/http');
            propData.push('Content-Transfer-Encoding: binary');
            propData.push('');

            propData.push('GET ' + getRequests[i].url + ' HTTP/1.1');
            propData.push('Accept: application/json;odata=verbose');
            propData.push('');
        }

        propData.push('--batch_' + batchBoundary + '--');

        const bodyBoundary = propData.join('\r\n');
        var requestHeaders = {
            'X-RequestDigest': this.digest,
            'Content-Type': `multipart/mixed; boundary="batch_${batchBoundary}"`
        };


        return this._fetch(this.baseURL + "_api/$batch", {
                method: "POST",
                headers: requestHeaders,
                credentials: "include",
                body: bodyBoundary
            })
            .then(function(raw) {
                if(typeof raw !== "string") {
                    return { ok: false, data: raw };
                }

                var responses = [],
                    parsedCtx = '',
                    parsedRaw = raw.split("\r\n\r\n");

                for(var p of parsedRaw) {
                    if(p.indexOf("No Content") > -1) {
                        if(p.indexOf("HTTP") === 0 && p.indexOf("204") > -1) {
                            responses.push({
                                ok: true,
                                data: true
                            });
                        } else {
                            responses.push({
                                ok: false,
                                data: false
                            });
                        }
                    } else if(p.indexOf("<") === 0) {
                        var parsedCtx = {};

                        try {
                            if(p.indexOf("<m:error") > -1) {
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
                    } else if(p.indexOf("{") === 0) {
                        try {
                            parsedCtx = JSON.parse(p.split("}\r\n")[0] + "}");
                            parsedCtx = {
                                ok: true,
                                data: parsedCtx
                            };
                        } catch (e) {
                            parsedCtx = {
                                ok: false,
                                data: p
                            };
                        }

                        responses.push(parsedCtx);
                    }
                }

                return responses;
            });
    },

    async _verifyDigest(refresh = false) {
        if(document.querySelector("#__REQUESTDIGEST") && document.querySelector("#__REQUESTDIGEST").value && !refresh) {
            this.digest = document.querySelector("#__REQUESTDIGEST").value;
        } else if(refresh || !this.digest) {
            const _this = this;
            return this._rest("_api/contextinfo", {
                    method: "POST"
                })
                .then(function(r) {
                    _this.digest = r?.GetContextWebInformation?.FormDigestValue || '';
                    return r;
                })
                .catch(function(e) {
                    _this._error(e);
                    return e;
                });
        }

        return this.digest;
    },

    async rest(url, fetchOptions = {}, config = {}) {
        await this._verifyDigest();
        this._headers = Object.assign({
            "Accept": "application/json; odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": this.digest
        }, this.headers);

        return this._rest(url, fetchOptions, config);
    },

    _rest(url, fetchOptions = {}, config = {}) {
        const _this = this;
        var urlFetch = _this.baseURL + url;
        return _this._fetch(urlFetch, fetchOptions)
            .then(function(resp) {
                if(resp?.d?.results) {
                    return resp.d.results.map(function(res) {
                        return _this._cleanObj(res, (config.select || []));
                    });
                } else if(resp?.d) {
                    return _this._cleanObj(resp.d, (config.select || []));
                }
                return resp;
            });
    },

    async _fetch(url, fetchOptions) {
        const _this = this;
        const controller = new AbortController();
        const abortFetch = setTimeout(function() {
            return controller.abort();
        }, _this.timeout);
        const urlFetch = url;

        const returnFetch = await fetch(urlFetch, Object.assign({
                    signal: controller.signal,
                    headers: _this._headers
                },
                fetchOptions
            ))
            .then(async function(resp) {
                var responseContent = false;
                try {
                    responseContent = await resp.text();
                } catch (e) {
                    if(resp.status > 199 && resp.status < 300) {
                        return true;
                    }
                    return false;
                }

                try {
                    return JSON.parse(responseContent);
                } catch (e) {
                    if(resp.status > 199 && resp.status < 300) {
                        return responseContent?.length ? responseContent : true;
                    }

                    _this._error("[SPHTTP] Internal error - The response was not a JSON.", [resp, fetchOptions]);
                    return false;
                }
            })
            .then(function(resp) {
                if(resp && resp["odata.error"] && resp["odata.error"]?.message?.value) {
                    throw (resp["odata.error"].message.value);
                } else if(resp?.error?.message?.value) {
                    throw (resp.error.message.value);
                } else if(fetchOptions?.method === "POST" && resp?.d?.Id) {
                    return resp.d;
                }
                return resp;
            });

        clearTimeout(abortFetch);
        return returnFetch;
    },

    _cleanObj(obj, listKeys = []) {
        if(this.cleanResponse && listKeys.length && listKeys[0] != "") {
            var newObject = {};
            if(listKeys.join().length > 0) {
                for(var key in listKeys) {
                    if(obj[listKeys[key]]) {
                        newObject[listKeys[key]] = obj[listKeys[key]];
                    } else if(obj[listKeys[key].split('/')[0]]) {
                        if(!newObject[listKeys[key].split('/')[0]]) {
                            newObject[listKeys[key].split('/')[0]] = [];
                        }

                        if(obj[listKeys[key].split('/')[0]].results && !newObject[listKeys[key].split('/')[0]].length) {
                            newObject[listKeys[key].split('/')[0]].push.apply(newObject[listKeys[key].split('/')[0]], obj[listKeys[key].split('/')[0]].results);
                        } else if(obj[listKeys[key].split('/')[0]] && obj[listKeys[key].split('/')[0]][listKeys[key].split('/')[1]]) {
                            newObject[listKeys[key].split('/')[0]][0] = obj[listKeys[key].split('/')[0]];
                        }
                    } else {
                        newObject[listKeys[key]] = typeof newObject[listKeys[key]] == "object" ? null : newObject[listKeys[key]];
                    }
                }
            }
            return newObject;
        }
        return obj;
    },

    _reset() {
        this.baseURL = "../";
        this.cleanResponse = true;
        this.digest = null;
        this.headers = {
            "Accept": "application/json; odata=verbose"
        };
        this._headers = {};
        this.timeout = 15000;
        this.top = 5000;
    },

    _error(e, obj = {}) {
        console.error(e, obj);
        return e;
    },

    _getFileBuffer(file) {
        return new Promise(function(resolve, reject) {
            try {
                var reader = new FileReader();
                reader.onload = function(e) {
                    resolve(e.target.result);
                }
                reader.onerror = function(e) {
                    reject(e.target.error);
                }
                reader.readAsArrayBuffer(file);
            } catch (error) {
                console.error('[ERROR] ' + error);
                reject(error);
            };
        });
    },

    _recursive(requestUrl, obj = {}) {
        var contentResponse = [];
        const _this = this;
        return new Promise(function(resolve, reject) {
            var content = [];
            var loop = async function(url) {
                _this._fetch(url).then(async function(resp) {
                        if(resp?.d?.results) {
                            content.push.apply(
                                content,
                                resp.d.results.map(function(res) {
                                    return _this._cleanObj(res, obj.select.split(","));
                                })
                            );
                        } else if(resp?.d) {
                            content.push.apply(content, _this._cleanObj(resp.d, obj.select.split(",")));
                        }

                        if(resp?.d?.__next) {
                            await loop(resp.d.__next);
                        } else {
                            resolve(content);
                        }
                    })
                    .catch(function(error) {
                        reject(error);
                    });
            }

            loop(requestUrl);
        });
    },
}