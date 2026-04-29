# Automatic filtering

## **Function Prototype**

```php
autoFilter(string $range): self
```

### **string $range**

> Filter/Filter Data Range

## Example

```php
$config = ['path' => './tests'];
$excel = new \Vtiful\Kernel\Excel($config);

$filePath = $excel->fileName("tutorial.xlsx")
     ->header(['name', 'age'])
     ->data([
         ['one', 10],
         ['two', 20],
         ['three', 30],
     ])
     ->autoFilter("A1:B3") // Add Filter/Filter
     ->output();
```