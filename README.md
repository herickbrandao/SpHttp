# SpHttp v0.1.2
A lightweight promise-based Javascript library for Sharepoint Rest services (7Kb ONLY!)

## Get Started
```html
<script type="text/javascript" src="SpHttp.min.js"></script>
<script type="text/javascript">
  // Get current user information
  console.log( 'The lib is working!', await SpHttp().user() );
</script>
```

## The Instance
```js
SpHttp({
  baseURL, // default: '../' (optional)
  headers, // default: { "Accept": "application/json; odata=verbose" } (optional)
  timeout  // default: 8000 (optional)
})
```

## Lists
```js
// Get List Items Example
SpHttp().list('ListName').get({
  top, // default: 5000 (optional)
  select, // example: ['ID','Title'] (optional)
  expand, // example: ['OtherList'] (optional)
  orderby, // example: 'Title asc' (optional)
  filter, // example: 'ID eq 1' (optional)
  recursive, // default: false (boolean) - get all list items [useful if the list is over 5000 items] (optional),
  ID, // example: 1 - returns a specific item (optional)
});

// Create List Item Example
SpHttp().list('ListName').post({
  Title: 'New Item' // List Info.
});

// Update List Item Example
SpHttp().list('ListName').put({
  ID: 1, // required
  Title: 'Update Item' // List Info.
});

// Delete List Item Example
SpHttp().list('ListName').del({
  ID: 1, // required
});

// Recycle List Item Example
SpHttp().list('ListName').recycle({
  ID: 1, // required
});

// Attachment List Example
SpHttp().list('ListName').attach({
  ID: 1, // required
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png' // filename that should be deleted (optional)
});
```

## Users
```js
// Get Current User Info.
SpHttp().user();

// Get User Info. by User Id
SpHttp().user(1);

// Get User Info. by Title
SpHttp().user('Name');

// Get User Groups by ID
SpHttp().user({
  ID: 1, // required
});
```

## Document Library
```js
SpHttp().attach({
  library: '/sites/myWebSite/Documents', // relative lib URL (required)
  name: 'filename.png', // filename for GET/POST/UPDATE requests (optional)
  startswith: true, // default: false - filters the file by name (optional)
  target: '#inputFile', // every file at <input type="file" id="inputFile" /> will be attached (optional)
  delete: 'filename.png' // filename that should be deleted (optional)
});
```

## Examples
Get List Item ID,Title By ID
```js
SpHttp().list('ListName').get({ ID: 11, select: ['ID', 'Title'] });
```

Get Lists over 5000 items
```js
SpHttp().list('ListName').get({ select: ['ID', 'Title'], recursive: true });
```

Attach File(s) at List Item
```html
<input type="file" id="inputFile" />

<script type="text/javascript" async>
  const attachs = await SpHttp().list('ListName').attach({ ID: 97 }); // getter

  document.getElementById("inputFile").addEventListener("change", function(e) {
    SpHttp().list('ListName').attach({ ID: 97, target: '#inputFile' }); // setter - Warning: this method does not overwrite!
  });
</script>
```

Document Library - Filter
```js
SpHttp().attach({ library: '/sites/myWebSite/Documents', name: 'file', startswith: true });
```

## Source of Inspiration
- [SpRestLib - Microsoft SharePoint REST JavaScript Library](https://github.com/gitbrent/SpRestLib/)
- [axios - Promise based HTTP client for the browser and node.js](https://github.com/axios/axios)
