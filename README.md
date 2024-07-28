# SpHttp v1.0.0-ALPHA
A lightweight promise-based Javascript library for Sharepoint Rest services

## Get Started
```html
<script type="text/javascript" src="sphttp.min.js"></script>
<script type="text/javascript">
  // Get current user information
  console.log( 'The lib is working!', await sphttp.user() );
</script>
```

## The Instance
```js
  sphttp.baseURL; // default: '../'
  sphttp.cleanResponse; // default: true
  sphttp.digest; // default: null (if it's empty, the code is going to set it)
  sphttp.headers; // default: { "Accept": "application/json; odata=verbose" }
  sphttp.timeout; // default: 15000
  sphttp.top; // default: 5000
  sphttp.version;
```

## Lists
```js
// Get List Items Example
sphttp.items('ListName', {
  top, // default: 5000 (optional)
  select, // example: ['ID','Title'] (optional)
  expand, // example: ['OtherList'] (optional)
  orderby, // example: 'Title asc' (optional)
  filter, // example: 'ID eq 1' (optional)
  recursive, // default: false (boolean) - get all list items [useful if the list is over 5000 items] (optional)
  ID, // example: 1 - returns a specific item (optional)
  versions, // boolean - if true, will return the item history (optional, needs ID)
  avoidcache: true, // default: false (optional)
});

// Create List Item Example
sphttp.add('ListName', {
  Title: 'New Item' // List Info.
});

// Update List Item Example
sphttp.update('ListName', {
  ID: 1, // required
  Title: 'Update Item' // List Info.
});

// Delete List Item Example
sphttp.delete('ListName', {
  ID: 1, // required
});

// Recycle List Item Example
sphttp.recycle('ListName', { ID: 1 });
// OR
sphttp.recycle('ListName', 1);

// Attachment List Example
sphttp.attach('ListName', {
  ID: 1, // required
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png' // filename that should be deleted (optional)
});
```

## Users
```js
// Get Current User Info.
sphttp.user();

// Get User Info. by User Id
sphttp.user(1);

// Get User Info. by Title
sphttp.user('Name');

// Get User Groups by ID
sphttp.user({ ID: 1 });
```

## Document Library
```js
sphttp.attachDoc({
  library: '/sites/myWebSite/Documents', // relative lib URL (required)
  name: 'filename.png', // filename for GET/POST/UPDATE requests (optional)
  startswith: true, // default: false - filters the file by name (optional)
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png', // filename that should be deleted (optional)
  recycle: false, // set true if you want (delete attribute required)
});
```

## Batch support
```js
await sphttp.batch([
    { url: "https://yourservice.sharepoint.com/sites/yoursite/_api/lists/getbytitle('listName')/items(75)", action: "UPDATE",
      item: {
        Title: 'UPDATE EXAMPLE', __metadata: {type: "SP.Data.listNameListItem"}
      }
    },
    { url: "https://yourservice.sharepoint.com/sites/yoursite/_api/lists/getbytitle('listName')/items", action: "POST",
      item: {
        Title: 'CREATE EXAMPLE', __metadata: {type: "SP.Data.listNameListItem"}
      }
    },
    { url: "https://yourservice.sharepoint.com/sites/yoursite/_api/lists/getbytitle('listName')/items?$select=Example", action: "GET" },
    { url: "https://yourservice.sharepoint.com/sites/yoursite/_api/lists/getbytitle('listName')/items?$select=ID", action: "GET" },
])
```

## Examples
Get List Item ID,Title By ID
```js
sphttp.items('ListName', { ID: 11, select: ['ID', 'Title'] });
```

Get Lists over 5000 items
```js
sphttp.items('ListName', { select: ['ID', 'Title'], recursive: true });
```

Make your own rest request
```js
sphttp.rest("_api/lists/getbytitle('ListName')/items?$skiptoken=Paged%3dTRUE%26p_ID%3d15000&$top=5000");
```

Get item history (versions)
```js
sphttp.items('ListName', { ID: 11, select: ['ID', 'Title'], versions: true });
```

Attach File(s) at List Item
```html
<input type="file" id="inputFile" />

<script type="text/javascript" async>
  const attachs = await sphttp.attach('ListName', { ID: 97 }); // getter

  document.getElementById("inputFile").addEventListener("change", function(e) {
    sphttp.attach('ListName', { ID: 97, target: '#inputFile' }); // setter - Warning: this method does not overwrite!
  });
</script>
```

Document Library - Filter
```js
sphttp.attachDoc({ library: '/sites/myWebSite/Documents', name: 'file', startswith: true });
```

## Source of Inspiration
- [SpRestLib - Microsoft SharePoint REST JavaScript Library](https://github.com/gitbrent/SpRestLib/)
- [axios - Promise based HTTP client for the browser and node.js](https://github.com/axios/axios)
