# VisDbDig - Visio Database Diagram

Tool to create Visio diagram of a database schema. This works in two parts. One command (VisDbDig.sql.exe) exports a SQL Server database schema to a set of JSON files. The second command (VisDbDig.exe) reads those files and generates a Visio diagram. This separation allows for creating other tools which will generate JSON files in the same format from other databases, APIs, documents, etc. To support doing this the types used to serialize the JSON files are in a separate library (VisDbDig.Model.dll).

## VisDbDig.sql.exe

Export the schema of a SQL Server DB to JSON files:

Usage

```
VisDbDig.sql.exe [OLEDB Connection String] [Output Directory]

Example:

VisDbDig.sql.exe "Provider=sqloledb;Data Source=.;Initial Catalog=WideWorldImporters;Integrated Security=SSPI;" "C:\temp\WideWorldImporters"
```

The command connects to the database specified in the [OLEDB connection string](https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/connection-string-syntax). It then writes the output to the specified directory. The directory is created if it doesn't exist. The following files are created:

### types.json

The tables and the fields in the files. E.g.

```json
{
    "Person" : [
        {
            "Name": "First_Name",
            "DataType": "nvarchar(128)",
            "OneToMany": false
        },
        {
            "Name": "Last_Name",
            "DataType": "nvarchar(128)",
            "OneToMany": false
        }
    ],
    "Test": [
        {
            "Name": "Score",
            "DataType": "int",
            "OneToMany": false
        }
    ]
}
```
This JSON file is deserialize into a `Dictionary<string,List<Field>>`'. The keys of the dictionary are the table names.

The OneToMany property on the Fields is always false when this tool exports a SQL Server schema. Exports form other data sources can set this property. This is then used to indicate the cardinally of relationships in the diagram.

### relationships.json

The relationships between tables. E.g.:

```json
[
    {
        "From": "Person",
        "To": "Tests",
    },
    {
        "From": "Person",
        "To": "Address",
    }
]
```

This JSON file is deserialize into a `List<Relationship>`.

### typenames.json

This file is simply a list of the types in the types.json file. E.g.:

```json
[
    "Person",
    "Test",
    "address"
]
```

### Other Database Types

Since this tool uses an OLEDB connection string it may work with other database types. I haven't tested this.

## VisDbDig.exe

This command creates a Visio diagram using the output of VisDbDig.sql.exe or another tool which creates the same file format.

### Usage 1
```
VisDbDig.exe [Input Directory]

Example:

VisDbDig.exe "C:\Temp\WideWorldImporters"
```

### Usage 2
```
VisDbDig.exe [Types JSON File] [Relationships JSON File] <Filter Types Text File>

Example:

VisDbDig.exe "C:\Temp\WideWorldImporters\Types.json" "C:\Temp\WideWorldImporters\Relationships.json" "C:\Temp\WideWorldImporters\Tables_To_Draw.txt"
```

The first usage is simple to use then you want to diagram the output of the VisDbDig.sql.exe tool. Simple specify the path to the folder containing the output of that command, e.g. the folder containing the `types.json` and `relationship.json` files.

Second usage allows for more control over the command. The paths to types and are relationship files a specified explicably. A optional list of tables to draw can be specified. This is a simple text file. If this file is provided then only tables whose names are in that file are drawn in the diagram. Otherwise all tables are drawn.

