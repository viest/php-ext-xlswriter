# Read file by data type

* The file is not supported for the `windows` system.
* Extended version is greater than or equal to `1.2.7`;

## Compiling

Add `--enable-reader` when compiling

```bash
./configure --enable-reader
```

##example

```bash
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial.xlsx')
     ->header(['Item', 'Cost'])
     ->output();

$excel->openFile('tutorial.xlsx')
     ->openSheet();

// When reading each row of cell data, you can specify each cell data type to read
Var_dump($excel->nextRow([
     \Vtiful\Kernel\Excel::TYPE_STRING, \Vtiful\Kernel\Excel::TYPE_STRING
]));
```