# SpHttp
A lightweight promise-based Javascript library for Sharepoint Rest services (4Kb Min ONLY!)

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
