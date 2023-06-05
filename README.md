# SpHttp v0.5.0
A lightweight promise-based Javascript library for Sharepoint Rest services (10Kb ONLY!)

## Get Started
```html
<script type="text/javascript" src="sphttp.min.js"></script>
<script type="text/javascript">
  // Get current user information
  console.log( 'The lib is working!', await sphttp().user() );
</script>
```

## The Instance
```js
sphttp({
  baseURL, // default: '../' (optional)
  headers, // default: { "Accept": "application/json; odata=verbose" } (optional)
  timeout  // default: 8000 (optional)
})
```

## Lists
```js
// Get List Items Example
sphttp().list('ListName').items({
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
sphttp().list('ListName').add({
  Title: 'New Item' // List Info.
});

// Update List Item Example
sphttp().list('ListName').update({
  ID: 1, // required
  Title: 'Update Item' // List Info.
});

// Delete List Item Example
sphttp().list('ListName').delete({
  ID: 1, // required
});

// Recycle List Item Example
sphttp().list('ListName').recycle({
  ID: 1, // required
});

// Attachment List Example
sphttp().list('ListName').attach({
  ID: 1, // required
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png' // filename that should be deleted (optional)
});

// Iterate List Example - Get a vast amount of data simultaneously (awesome for large lists like 50k of items)
sphttp().list('ListName').iterate({
  top, // default: 5000 (optional)
  select, // example: ['ID','Title'] (optional)
  expand, // example: ['OtherList'] (optional)
  action, // bind a function after each request, example: a => { return a.filter(b => b.ID === 93) } (optional)
});
```

## Users
```js
// Get Current User Info.
sphttp().user();

// Get User Info. by User Id
sphttp().user(1);

// Get User Info. by Title
sphttp().user('Name');

// Get User Groups by ID
sphttp().user({
  ID: 1, // required
});
```

## Document Library
```js
sphttp().attach({
  library: '/sites/myWebSite/Documents', // relative lib URL (required)
  name: 'filename.png', // filename for GET/POST/UPDATE requests (optional)
  startswith: true, // default: false - filters the file by name (optional)
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png', // filename that should be deleted (optional)
  recycle: false, // set true if you want (delete attribute required)
});
```

## Extending library
```js
// method arguments
sphttp().extend(name, callback);

// example
sphttp().extend('exampleItems', function(listName) {
  return this.list(listName).items();
});

// then
await sphttp().exampleItems('myExampleList');
```

## Examples
Get List Item ID,Title By ID
```js
sphttp().list('ListName').get({ ID: 11, select: ['ID', 'Title'] });
```

Get Lists over 5000 items
```js
sphttp().list('ListName').get({ select: ['ID', 'Title'], recursive: true });
```

Make your own rest request
```js
sphttp().rest("_api/lists/getbytitle('ListName')/items?$skiptoken=Paged%3dTRUE%26p_ID%3d15000&$top=5000");
```

Get item history (versions)
```js
sphttp().list('ListName').get({ ID: 11, select: ['ID', 'Title'], versions: true });
```

Attach File(s) at List Item
```html
<input type="file" id="inputFile" />

<script type="text/javascript" async>
  const attachs = await sphttp().list('ListName').attach({ ID: 97 }); // getter

  document.getElementById("inputFile").addEventListener("change", function(e) {
    sphttp().list('ListName').attach({ ID: 97, target: '#inputFile' }); // setter - Warning: this method does not overwrite!
  });
</script>
```

Document Library - Filter
```js
sphttp().attach({ library: '/sites/myWebSite/Documents', name: 'file', startswith: true });
```

Get App Context (can be used for token refresh)
```js
await sphttp().contextinfo();
```

## Source of Inspiration
- [SpRestLib - Microsoft SharePoint REST JavaScript Library](https://github.com/gitbrent/SpRestLib/)
- [axios - Promise based HTTP client for the browser and node.js](https://github.com/axios/axios)
