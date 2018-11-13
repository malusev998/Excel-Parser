# Excel Converter (NodeJS)

> Simple but powerful excel to jsobject converter


## Usage

```javascript
	let excel = new Excel('path/to/excel/file' ); // Instanciate Excel
	
	// To get object from worksheet
	// Pass the number for worksheet ( starts from 1, not from 0 )
	try {
		// Object will assume that the first row is header with names
		let object = excel.getWorksheet( 1 ); //  { header: firstRow, body: restOfRows }	
	}
	catch( err ) {
		// Error will accure if excel worksheet doesn't exist or mapping has failed
		console.log( err );
	}
```


## Upcoming features

1. Get  by excel range ( A1:B3 )

1. Get all rows

1. Get Indiviual cell

1. Insert into excel

1. Update fields


## Dependencies 

1. AdmZip ( adm-zip )

1. XMLDOM

## Licence MIT
