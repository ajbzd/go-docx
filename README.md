# docx
## Introduction
docx is a simple library to creating DOCX file in Go.

## Getting Started
### Install
Go modules supported
```sh
go get github.com/ajbzd/go-docx
```
Import:
```sh
import "github.com/ajbzd/go-docx"
```

### Usage
**Example:**
```go
package main

import (
	"github.com/ajbzd/go-docx"
)

func main() {
	f := docx.NewFile()
	// add new paragraph
	para := f.AddParagraph()
	// add text
	para.AddText("test")

	para.AddText("test font size").Size(22)
	para.AddText("test color").Color("808080")
	para.AddText("test font size and color").Size(22).Color("121212")

	nextPara := f.AddParagraph()
	nextPara.AddLink("google", `http://google.com`)

	f.Save("./test.docx")
}

```
