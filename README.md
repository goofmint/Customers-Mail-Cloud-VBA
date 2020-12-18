# Customers Mail Cloud library for VBA

## Dependence

- [VBA-tools/VBA-JSON: JSON conversion and parsing for VBA](https://github.com/VBA-tools/VBA-JSON)
- Microsoft Script Runtime

## Usage

### Init

```vb
Dim client As clsCustomersMailCloud
Set client = New clsCustomersMailCloud
client.ApiUser = "API_USER"
client.ApiKey = "API_KEY"
client.Trial
```

### Set to address

```vb
Dim toAddress As clsCustomersMailCloudAddress: Set toAddress = New clsCustomersMailCloudAddress
toAddress.Name = "Tester"
toAddress.Address = "tester1@smtps.jp"
client.AddTo toAddress

Dim toAddress2 As clsCustomersMailCloudAddress: Set toAddress2 = New clsCustomersMailCloudAddress
toAddress2.Name = "Tester 2"
toAddress2.Address = "tester2@smtps.jp"
client.AddTo toAddress2
```

### Set from address

```vb
Dim fromAddress As clsCustomersMailCloudAddress: Set fromAddress = New clsCustomersMailCloudAddress
fromAddress.Name = "Admin"
fromAddress.Address = "info@smtps.jp"
client.SetFrom fromAddress
```

### Set Subject and Mail body

```vb
client.Subject = "Test mail"
client.Text = "Mail body"
```

### Send Mail!

```vb
client.Send
```

## License

MIT
