sphttp._attachDoc = async function(config = {}) {
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
    var _this = this;
    var appends = [];

    for(var i in items) {
        if(typeof items[i] === 'object') {
            appends.push(this._getFileBuffer(items[i]));
        }
    }

    return Promise.all(appends).then(async function(promises) {
        var httpRes = [],
            url = '';

        for(var i in promises) {
            _this._headers = Object.assign({
                "content-length": promises[i].byteLength
            }, _this._headers);

            var bytes = new Uint8Array(promises[i]);
            var name = config.name ? config.name : items[i].name;
            name = config.prefix ? config.prefix + name : name;
            url = "_api/web/GetFolderByServerRelativeUrl('" + config.library + "')/Files/Add(url='" + name + "', overwrite=true)?$expand=ListItemAllFields&$select=Id";

            var ctx = await _this._rest(url, {
                method: "POST",
                body: promises[i]
            });

            httpRes.push(ctx);
        }

        for(let i in httpRes) {
            if(httpRes[i] && httpRes[i].d) {
                httpRes[i] = httpRes[i].d;
            }
        }

        return httpRes;
    });
};