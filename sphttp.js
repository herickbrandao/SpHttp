var sphttp = (function(options = {}) {

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
                if(resp&&resp["odata.error"]&&resp["odata.error"].message&&resp["odata.error"].message.value) {
                    throw(resp["odata.error"].message.value);
                }
                else if(resp&&resp["error"]&&resp["error"].message&&resp["error"].message.value) {
                    throw(resp["error"].message.value);
                } else if(fetchOptions&&fetchOptions.method==="POST"&&resp&&resp.d&&resp.d.Id) {
                    return resp.d;
                }
                return resp;
            });
        clearTimeout(id);
        return response;
    }

    function cleanObj(obj, allowNulls = true) {
        var newObject = {};
        if(listKeys.join().length>0) {
            for(var key in listKeys) {
                if(obj[listKeys[key]]) {
                    newObject[listKeys[key]] = obj[listKeys[key]];
                } else if(obj[listKeys[key].split('/')[0]]) {
                    if(!newObject[listKeys[key].split('/')[0]]) { newObject[listKeys[key].split('/')[0]] = []; }
                    if(obj[listKeys[key].split('/')[0]].results&&!newObject[listKeys[key].split('/')[0]].length) {
                        newObject[listKeys[key].split('/')[0]].push.apply(newObject[listKeys[key].split('/')[0]],obj[listKeys[key].split('/')[0]].results);
                    } else if(obj[listKeys[key].split('/')[0]]&&obj[listKeys[key].split('/')[0]][listKeys[key].split('/')[1]]) {
                        newObject[listKeys[key].split('/')[0]][0] = obj[listKeys[key].split('/')[0]];
                    }
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
            recycle: recycleList,
            attach: attachList,
            iterate: getIterate
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
            if(config.versions) { url+= '/versions'; }
            if(config.select&&config.select.length>0) { url += first+'$select='+config.select; first = '&'; }
            if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; }
            if(config.avoidcache) { url += first+'$v='+window.crypto.randomUUID(); first = '&'; }
            return rest(url);
        }

        url += '?$top='+config.top;
        first = '&';
        if(config.select&&config.select.length>0) { url += first+'$select='+config.select; first = '&'; }
        if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; }
        if(config.filter&&config.filter.length>0) { url += first+'$filter='+config.filter; first = '&'; }
        if(config.avoidcache) { url += first+'$v='+window.crypto.randomUUID(); first = '&'; }

        if(!config.recursive) {
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
        options.headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value
        },options.headers);

        var url = "_api/lists/getbytitle('"+listName+"')/items";
        return fetchWithTimeout(url, { method: "POST", body: typeof item === 'object' ? JSON.stringify(item) : item });
    }

    function putList(item) {
        options.headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
            "IF-MATCH": "*",  
            "X-HTTP-Method": "MERGE"
        },options.headers);

        var url = "_api/lists/getbytitle('"+listName+"')/items";

        if(typeof item === 'object' && item.ID) {
            url += '('+item.ID+')';
        } else {
            throw('[ERROR] Put request data is wrong!');
        }

        return fetchWithTimeout(url, { method: "POST", body: typeof item === 'object' ? JSON.stringify(item) : item });
    }

    function recycleList(item) {
        options.headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
        },options.headers);

        var url = "_api/lists/getbytitle('"+listName+"')/items";

        if(typeof item === 'object' && item.ID) {
            url += '('+item.ID+')/recycle()';
        } else {
            throw('[ERROR] Recycle request data is wrong!');
        }

        return fetchWithTimeout(url, { method: "POST", body: typeof item === 'object' ? JSON.stringify(item) : item });
    }

    function delList(item) {
        options.headers = Object.assign({
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
            "IF-MATCH": "*",  
            "X-HTTP-Method": "DELETE"
        },options.headers);

        var url = "_api/lists/getbytitle('"+listName+"')/items";

        if(typeof item === 'object' && item.ID) {
            url += '('+item.ID+')';
        } else if(typeof item === 'number') {
            url += '('+item+')';
        } else {
            throw('[ERROR] Delete request data is wrong!');
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

        throw('[ERROR] User Request Failed!');
    }

    function attachList(config = {}) {
        if(!window.FileReader) { throw('[ERROR] Your browser does not have support for FileReader!'); }
        if(isNaN(config.ID)) { throw('[ERROR] List ID not found at attach request!'); }
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
                options.headers = Object.assign({
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "X-HTTP-Method": "DELETE",
                },options.headers);
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
        if(!window.FileReader) { throw('[ERROR] Your browser does not have support for FileReader!'); }
        if(!config.library) { throw('[ERROR] List library not found at attach request!'); }
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
                options.headers = Object.assign({
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "content-length": promises[i].byteLength
                },options.headers);
                var bytes = new Uint8Array(promises[i]);
                var name = config.name ? config.name : items[i].name;
                url = "_api/web/GetFolderByServerRelativeUrl('"+config.library+"')/Files/Add(url='"+name+"', overwrite=true)";
                
                httpRes.push(fetchWithTimeout(url, { method: "POST", body: promises[i] }));
            }

            if(typeof config.delete === "string" && (config.delete&&config.delete.length>0)) {
                options = JSON.parse(optbkp); // options reset
                options.headers = Object.assign({
                    "X-RequestDigest": document.querySelector("#__REQUESTDIGEST").value,
                    "X-HTTP-Method": "DELETE",
                },options.headers);
                
                url = "_api/web/getfilebyserverrelativeurl('"+config.library+'/'+config.delete+"')";
                if(config.recycle&&config.recycle==true) { url += '/recycle()'; }
                return rest(url, { method: "POST" });
            } else if(config.delete) {
                throw('[ERROR] Filename is missing!');
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
                    reject(e.target.error);
                }
                reader.readAsArrayBuffer(file);
            } catch(error) {
                console.error('[ERROR] '+error);
                reject(error);
            };
        });
    }

    async function getIterate(config =  {}) {
        if(!config.top) { config.top = 5000; }
        if(config.select&&!config.select.includes('ID')) { config.select.push('ID'); }
        if(!config.select) { config.select = []; }
        if(!config.expand) { config.expand = []; }
        if(typeof config.select === 'object') { config.select = config.select.join(); }
        if(typeof config.expand === 'object') { config.expand = config.expand.join(); }

        var scope = [];
        scope[0] = rest("_api/lists/getbytitle('"+listName+"')/items?$top=1&$select=Id&$orderby=Id asc");
        scope[1] = rest("_api/lists/getbytitle('"+listName+"')/items?$top=1&$select=Id&$orderby=Id desc");
        var total = await Promise.all(scope);
        config.total = [];
        config.total[0] = total[0][0]&&total[0][0].Id ? total[0][0].Id : 0;
        config.total[1] = total[1][0]&&total[1][0].Id ? total[1][0].Id : 0;
        config.total[2] = (config.total[1] - config.total[0]) + 1;
        config.total[3] = parseInt( config.total[2] / config.top ) + 1;
        
        listKeys = config.select.split(',');
        var url = '?$top='+config.top;
        if(config.select&&config.select.length>0) { url += '&$select='+config.select; }
        if(config.expand&&config.expand.length>0) { url += '&$expand='+config.expand; }

        var promises = [], results = [];
        config.total[0] -= 1;
        for(var i = 0; i < config.total[3]; i++) {
            var info = "_api/web/lists/getbytitle('"+listName+"')/items"+url;

            if(i>0) {
                config.total[0] += config.top;
                info += "&$skiptoken=Paged%3dTRUE%26p_ID%3d"+config.total[0];
            }

            var iterate;
            if(config.action && typeof config.action === 'function') {
                iterate = rest(info).then(config.action);
            } else {
                iterate = rest(info);
            }

            promises.push( iterate );
        }

        var resp = await Promise.all(promises);
        const uniqueIds = new Set();
        var final = [];

        for(var res of resp) {
            if(res.error) {
                final = res;
                break;
            }
            var r = res.filter(function(a) {
                const isDuplicate = uniqueIds.has(a.ID);
                uniqueIds.add(a.ID);
                return !isDuplicate;
            });
            final.push.apply(final,r);
        }

        return final;
    }

    function contextinfo() {
        return rest('_api/contextinfo', { method: "POST" });
    }

    function extend(name,callback) {
        window.__sphttp_extensions[name] = callback;
    }

    window.__sphttp_extensions = window.__sphttp_extensions || {};

    return {
        list,
        user,
        attach,
        rest,
        contextinfo,
        extend,
        fetch: fetchWithTimeout,
        options,
        version: '0.5.0',
        ...window.__sphttp_extensions
    };
});