# 下载

## 示例

```php
function getTmpDir(): string
{
    $tmp = ini_get('upload_tmp_dir');

    if ($tmp !== False && file_exists($tmp)) {
        return realpath($tmp);
    }

    return realpath(sys_get_temp_dir());
}

$config = [
    'path' => getTmpDir() . '/',
];

$fileName   = 'tutorial01.xlsx';
$xlsxObject = new \Vtiful\Kernel\Excel($config);

// Init File
$fileObject = $xlsxObject->fileName($fileName);

// Writing data to a file ......

// Outptu
$filePath = $fileObject->output();

// Set Header
header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header('Content-Disposition: attachment;filename="' . $fileName . '"');
header('Content-Length: ' . filesize($filePath));
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate');
header('Cache-Control: max-age=0');
header('Pragma: public');

if (copy($filePath, 'php://output') === false) {
    // Throw exception
}

// Delete temporary file
@unlink($filePath);
```

