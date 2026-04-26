# Sheet protection

Enable Excel's worksheet protection to prevent users from editing cells. Protection can be applied with or without a password, and can be selectively lifted on individual cells, rows, or columns by combining it with the `Format::unlocked()` style.

## Function Prototype

```php
protection(?string $password = null): self
```

### **string $password**

> Optional. The unlock password. When omitted the sheet is still protected, but anyone can unlock it from the Excel UI without a password prompt.

Topics:

* [Password protection](password.md)
* [Unlock cells](unlock.md)
