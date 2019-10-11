# Freeze cell

## **Function Prototype**

```php
freezePanes(int $row, int $column): self
```

### **int $row**

> Line number

### **int $column**

> column number

##example

```php
freezePanes(1, 0); // freeze the first line
freezePanes(0, 1); // Freeze the first column
freezePanes(1, 1); // Freeze the first row and first column
```