# Docx
Docx is a simple package for for filling out documents prepared in [Office Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) format,
usually saved as a filename with `.docx` extension.

It looks for variables in text and replaces them with defined values.

You can check [example template document](go-docx-test.docx).

# Installation

```
go get -u github.com/elblox/go-docx
```

# Usage

```go
	dict := map[string]string{
		"[variable1]": "Some text",
		"[variable2]": "Some other text",
	}
	output, err := os.Create("output.docx")
	if err != nil {
		return err
	}
	input, err := os.Open("input.docx")
	if err != nil {
		return err
	}
	stat, err := input.Stat()
	if err != nil {
		return err
	}
	// we need file size because of zip.Reader
	_, err = docx.New(input, stat.Size()).
		Replace(dict).
		WriteTo(output)
	if err != nil {
		return err
	}
```

You can also check [docx_test.go](docx_test.go).
