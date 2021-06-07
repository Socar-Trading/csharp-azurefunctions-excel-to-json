# azurefunctions-excel-to-json

Sample Azure Function to convert Excel (XLS or XLSX) files and CSV to JSON strings.

## Use cases

A few use cases this solution can be used on:

- Data Science projects, to convert different sources into a common data format.
- To convert user inputed data from Excel files into a developer friendly format as JSON.

## Sending a request directly

You can send a **POST** request passing the file as a [FormData](https://developer.mozilla.org/pt-BR/docs/Web/API/FormData) to the Azure Functions endpoint.

**POST**

```JSON
{
    "data": dataBlobContent
}
```

## Sending from a HTML form

You can send a **POST** request from a web page using a form as input.

*This example uses jQuery, but you can do this from native javascript, React, Angular and any other web framework.*

```JS
    // Gets the first file uploaded from the form component.
    var data = new FormData();
    data.append("file",$('<your-form-id>')[0].files[0]);
    
    // AJAX request.
    $.ajax({
        method: "POST",
        url: "<your-endpoint-url>",
        crossDomain: true,
        enctype: 'multipart/form-data',
        cache : false,                  // disables caching for this request
        processData : false,            // ensures no processing will be done on the content by the jQuery library
        data: data                      // the content as a FormData file.
    });
```

