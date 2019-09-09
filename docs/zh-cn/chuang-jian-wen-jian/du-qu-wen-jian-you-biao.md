# 读取文件（游标）

* 读取文件暂不支持 `windows` 系统；
* 扩展版本大于等于 `1.2.7`；

### 编译

编译时需添加 `--enable-reader` 

```bash
./configure --enable-reader
```

### 示例

```bash
$config   = ['path' => './tests'];
$excel    = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName('tutorial.xlsx')
    ->header(['Item', 'Cost'])
    ->output();
    
$excel->openFile('tutorial.xlsx')
    ->openSheet();

var_dump($excel->nextRow()); // ['Item', 'Cost']
var_dump($excel->nextRow()); // NULL
```



