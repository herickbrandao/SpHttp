const SpHttp = (function(options = {}) {

    if(!options.baseURL) { options.baseURL = '../'; }
    if(!options.headers) { options.headers = { "Accept": "application/json; odata=verbose" }; }
    if(!options.timeout) { options.timeout = 8000; }
    options.headers = Object.assign({ "Accept": "application/json; odata=verbose" }, options.headers);

    var listName = '', listKeys = [], userName = {};

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
            .then(function(resp) { return options.headers["X-HTTP-Method"] ? true : resp.json(); })
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
        for(key in listKeys) {
            if(obj[listKeys[key]]) {
                newObject[listKeys[key]] = obj[listKeys[key]];
            } else if(allowNulls) {
                newObject[listKeys[key]] = null;
            }
        }
        return listKeys.length>0 ? newObject : obj;
    }

    function list(name) {
        listName = name;
        return {
            get: getList,
            post: postList,
            put: putList,
            update: putList,
            del: delList,
            delete: delList,
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
            if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; }
            return rest(url);
        }

        url += '?$top='+config.top;
        first = '&';
        if(config.select&&config.select.length>0) { url += first+'$select='+config.select; first = '&'; }
        if(config.expand&&config.expand.length>0) { url += first+'$expand='+config.expand; first = '&'; }

        if(!config.recursive) {
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
        switch(typeof config) {
            case "string":
                return rest("/_api/web/siteusers?$filter=startswith(Title,'"+config+"')");
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

    return {
        list,
        user,
        rest
    };
});