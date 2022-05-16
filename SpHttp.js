const SpHttp = (function(options = {}) {

    if(!options.baseURL) { options.baseURL = '../'; }
    if(!options.headers) { options.headers = { "Accept": "application/json; odata=verbose" }; }
    if(!options.timeout) { options.timeout = 8000; }
    options.headers = Object.assign({ "Accept": "application/json; odata=verbose" }, options.headers);

    var listName = '', listKeys = [], optbkp = JSON.stringify(options);

    function rest(url, config = {}) {
        return fetchWithTimeout(url, config).then(function(resp) {
            if(resp&&resp.d&&resp.d.results) { return resp.d.results.map(function(res) { return cleanObj(res); }); }
            else if(resp&&resp.d) { return cleanObj(resp.d); }
            return resp;
        });
    }

    async function fetchWithTimeout(url, fetchOptions, extra = false) {
        if(typeof extra === 'object' && extra.cleanURL) { options.baseURL = ''; }
        const controller = new AbortController();
        const id = setTimeout(function() { return controller.abort(); }, options.timeout);
        const response = await fetch(options.baseURL+url, Object.assign({ signal: controller.signal, headers: options.headers }, fetchOptions))
            .then(async function(resp) { var retorno = true; try { retorno = await resp.json(); } catch(e) { return true; } return retorno; })
            .then(function(resp) {
                if(resp&&resp["odata.error"]&&resp["odata.error"].message&&resp["odata.error"].message.value) { return { error: resp["odata.error"].message.value }; }
                else if(resp&&resp["error"]&&resp["error"].message&&resp["error"].message.value) { return { error: resp["error"].message.value }; }
                return resp;
            })
            .catch(function(error) { return { error: { message: error } }; });
        clearTimeout(id);
        return response;
    }

    function cleanObj(obj, allowNulls = true) {
        var newObject = {};
        if(listKeys.join().length>0) {
            for(var key in listKeys) {
                if(obj[listKeys[key]]) {
                    newObject[listKeys[key]] = obj[listKeys[key]];
                } else if(allowNulls) {
                    newObject[listKeys[key]] = null;
                }
            }
        }
        return listKeys.join().length>0 ? newObject : obj;
    }

    function list(name) {
        listName = name;
        options = JSON.parse(optbkp); // options reset

        return {
            get: getList,
            items: getList,
            add: postList,
            post: postList,
            create: postList,
            put: putList,
            update: putList,
            del: delList,
            delete: delList,
            attach: attachList
        };
    }

    function getList(config =  {}) {
        if(!config.top) { config.top = 5000; }
        if(!config.select) { config.select = []; }
        if(!config.expand) { config.expand = []; }
        if(typeof config.select === 'object') { config.select = config.select.join(); }
        if(typeof config.expand === 'object') { config.expand = config.expand.join(); }

        listKeys = config.select.split(',');
        var url = "_api/lists/getbytitle('"+listName+"')/items";
        var first = '?';

        if(config.ID) {
            url += '('+config.ID+')';
            if(config.select&&config.select.length>0) { url += first+'$select='+config.select; first = '&'; }
            if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; listKeys = ['']; }
            return rest(url);
        }

        url += '?$top='+config.top;
        first = '&';
        if(config.select&&config.select.length>0) { url += first+'$select='+config.select; first = '&'; }
        if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; listKeys = ['']; }

        if(!config.recursive) {
            if(config.filter&&config.filter.length>0) { url += first+'$filter='+config.filter; first = '&'; }
            if(config.orderby&&config.orderby.length>0) { url += first+'$orderby='+config.orderby; first = '&'; }
            return rest(url);
        }

        url = options.baseURL + url;
        var content = [];
        return new Promise(function(resolve,reject) {
            
            
            async function loop(url) {
                fetchWithTimeout(url,{},{ cleanURL: true }).then(async function(resp) {
                    if(resp&&resp.d&&resp.d.results) { content.push.apply(content, resp.d.results.map(function(res) { return cleanObj(res); })); }
                    else if(resp&&resp.d) { content.push.apply(content,cleanObj(resp.d)); }

                    if(resp&&resp.d&&resp.d.__next) { await loop(resp.d.__next); }
                    else { resolve(content);}
                })
                .catch(function(error) { reject(error); });
            }

            loop(url);
        }).then(function(res) { return content; });
    }

    function postList(item) {
        options.headers = Object.assign(options.headers, {
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
        });

        var url = "_api/lists/getbytitle('"+listName+"')/items";
        return fetchWithTimeout(url, { method: "POST", body: typeof item === 'object' ? JSON.stringify(item) : item });
    }

    function putList(item) {
        options.headers = Object.assign(options.headers, {
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
            "IF-MATCH": "*",  
            "X-HTTP-Method": "MERGE"
        });

        var url = "_api/lists/getbytitle('"+listName+"')/items";

        if(typeof item === 'object' && item.ID) {
            url += '('+item.ID+')';
        } else {
            console.error('[ERROR] Put request data is wrong!');
            return { error: '[ERROR] Put request data is wrong!' };
        }

        return fetchWithTimeout(url, { method: "POST", body: typeof item === 'object' ? JSON.stringify(item) : item });
    }

    function delList(item) {
        options.headers = Object.assign(options.headers, {
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
            "IF-MATCH": "*",  
            "X-HTTP-Method": "DELETE"
        });

        var url = "_api/lists/getbytitle('"+listName+"')/items";

        if(typeof item === 'object' && item.ID) {
            url += '('+item.ID+')';
        } else if(typeof item === 'number') {
            url += '('+item+')';
        } else {
            console.error('[ERROR] Delete request data is wrong!');
            return { error: '[ERROR] Delete request data is wrong!' };
        }

        return fetchWithTimeout(url, { method: "POST" });
    }

    function user(config = false) {
        options = JSON.parse(optbkp); // options reset

        switch(typeof config) {
            case "string":
                var searches = [
                    config.toLowerCase(),
                    config.toUpperCase(),
                ];
                
                // Capitalize
                const str = searches[0];
                const arr = str.split(" ");
                for (var i = 0; i < arr.length; i++) { arr[i] = arr[i].charAt(0).toUpperCase() + arr[i].slice(1); }
                searches[2] = arr.join(" ");

                // Normalize
                searches[3] = config.normalize('NFD').replace(/[\u0300-\u036f]/g, "");
                
                var url = '_api/web/siteusers?$filter=(';
                url += "(startswith(Title,'"+searches[0]+"')) or ";
                url += "(startswith(Title,'"+searches[1]+"')) or ";
                url += "(startswith(Title,'"+searches[3]+"')) or ";
                url += "(startswith(Title,'"+searches[2]+"')))";
                return rest(url);
                break;
            case "number":
                return rest('_api/Web/SiteUserInfoList/Items('+config+')');
                break;
            case "object":
                if(config.ID) { return rest("_api/web/GetUserById(" + config.ID + ")/Groups"); }
                break;
            case "boolean":
                return rest('_api/Web/CurrentUser?$expand=groups');
                break;
        }

        console.error('[ERROR] User Request Failed!');
        return { error: '[ERROR] User Request Failed!' };
    }

    function attachList(config = {}) {
        if(!window.FileReader) { return { error: '[ERROR] Your browser does not have support for FileReader!' }; }
        if(isNaN(config.ID)) { return { error: '[ERROR] List ID not found at attach request!' }; }
        if(!config.target&&!config.delete) { return rest("_api/lists/getbytitle('"+listName+"')/items("+config.ID+")/AttachmentFiles"); }

        var items = typeof config.target==="string"&&document.querySelector(config.target) ? document.querySelector(config.target).files : false;
        var appends = [];

        for(var i in items) { if(typeof items[i] === 'object') { appends.push(getFileBuffer(items[i])); } }

        return Promise.all(appends).then(function(promises) {
            var httpRes = [], url = '';

            for(var i in promises) {
                options.headers = Object.assign(options.headers, {
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "content-length": promises[i].byteLength
                });
                var bytes = new Uint8Array(promises[i]);
                var name = items[i].name;
                url = "_api/lists/GetByTitle('" + listName + "')/items(" + config.ID + ")/AttachmentFiles/add(FileName='" + name + "')";
                
                httpRes.push(fetchWithTimeout(url, { method: "POST", body: promises[i] }));
            }

            if(config.delete) {
                options = JSON.parse(optbkp); // options reset
                options.headers = Object.assign(options.headers, {
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "X-HTTP-Method": "DELETE",
                });
                url = "_api/lists/GetByTitle('" + listName + "')/items(" + config.ID + ")/AttachmentFiles/getByFileName('" + config.delete + "')";
                return rest(url, { method: "DELETE" });
            }

            return Promise.all(httpRes).then(function(res) { 
                for(var i in res) { if(res[i]&&res[i].d) { res[i] = res[i].d; } }
                return res;
            });
        });  
    }

    function attach(config = {}) {
        if(!window.FileReader) { return { error: '[ERROR] Your browser does not have support for FileReader!' }; }
        if(!config.library) { return { error: '[ERROR] List library not found at attach request!' }; }
        if(!config.name) { config.tname = ''; } else if(!config.startswith) { config.tname = "('"+config.name+"')"; }
        if(!config.target&&!config.delete) {
            if(config.startswith&&config.name) { config.tname = "?$filter=startswith(Name,'"+config.name+"')"; }
            return rest("_api/web/GetFolderByServerRelativeUrl('"+config.library+"')/Files"+config.tname);
        }

        var items = typeof config.target==="string"&&document.querySelector(config.target) ? document.querySelector(config.target).files : false;
        var appends = [];

        for(var i in items) { if(typeof items[i] === 'object') { appends.push(getFileBuffer(items[i])); } }

        return Promise.all(appends).then(function(promises) {
            var httpRes = [], url = '';

            for(var i in promises) {
                options.headers = Object.assign(options.headers, {
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "content-length": promises[i].byteLength
                });
                var bytes = new Uint8Array(promises[i]);
                var name = config.name ? config.name : items[i].name;
                url = "_api/web/GetFolderByServerRelativeUrl('"+config.library+"')/Files/Add(url='"+name+"', overwrite=true)";
                
                httpRes.push(fetchWithTimeout(url, { method: "POST", body: promises[i] }));
            }

            if(typeof config.delete === "string" && (config.delete&&config.delete.length>0)) {
                options = JSON.parse(optbkp); // options reset
                options.headers = Object.assign(options.headers, {
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "X-HTTP-Method": "DELETE",
                });
                
                url = "_api/web/GetFolderByServerRelativeUrl('"+config.library+"')/Files"+config.delete;
                return rest(url, { method: "DELETE" });
            } else if(config.delete) {
                return { error: '[ERROR] Filename is missing!' };
            }

            return Promise.all(httpRes).then(function(res) { 
                for(var i in res) { if(res[i]&&res[i].d) { res[i] = res[i].d; } }
                return res;
            });
        });  
    }

    function getFileBuffer(file) {
        return new Promise(function(resolve,reject) {
            try {
                var reader = new FileReader();
                reader.onload = function(e) {
                    resolve(e.target.result);
                }
                reader.onerror = function(e) {
                    reject({ error: e.target.error });
                }
                reader.readAsArrayBuffer(file);
            } catch(error) {
                console.error('[ERROR] '+error);
                reject({ error: error });
            };
        });
    }

    return {
        list,
        user,
        attach,
        version: '0.1.1'
    };
});