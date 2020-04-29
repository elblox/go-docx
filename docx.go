package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

const documentXML = "word/document.xml"

// Docx can manipulate .docx files created by MS Word 2007+
type Docx struct {
	zipReader      *zip.Reader
	err            error
	dict           Dict
	openingBracket rune
	closingBracket rune
}

// Dict is a dictionary with variables and values to which they should be replaced
type Dict map[string]string

// New creates Docx instance
func New(r io.ReaderAt, size int64) *Docx {
	doc := new(Docx)
	doc.zipReader, doc.err = zip.NewReader(r, size)
	doc.openingBracket = '['
	doc.closingBracket = ']'
	return doc
}

// Brackets allows to configure characters which symbolice start and end of variable
func (doc *Docx) Brackets(opening, closing rune) *Docx {
	doc.openingBracket = opening
	doc.closingBracket = closing
	return doc
}

// Replace stores dictionary of words to replace
func (doc *Docx) Replace(dict map[string]string) *Docx {
	doc.dict = dict
	return doc
}

// Buffer is a slice of XML tokes which are buffered before saving them in a file
type Buffer []xml.Token

// Flush saves all tokens to XML file and cleans the buffer
func (buffer *Buffer) Flush(encoder *xml.Encoder) error {
	for _, token := range *buffer {
		err := encoder.EncodeToken(fixNS(token))
		if err != nil {
			return err
		}
	}
	buffer.Clean()
	return nil
}

// Clean removes tokens from a buffer and keeps capacity untouched
func (buffer *Buffer) Clean() {
	*buffer = (*buffer)[:0]
}

// Process converts CharData tokens from a buffer to one string
// and replaces variables with values from a dictionary
func (buffer *Buffer) Process(encoder *xml.Encoder, dict Dict) error {
	varName := ""
	// wt indicates if we are currently in <w:t> XML element (where text is stored)
	// all non-wt elements should be ignored when extracting a variable name
	wt := true
	for _, token := range *buffer {
		if start, ok := token.(xml.StartElement); ok && isWT(start.Name) {
			wt = true
		}
		if end, ok := token.(xml.EndElement); ok && isWT(end.Name) {
			wt = false
		}
		if charData, ok := token.(xml.CharData); ok && wt {
			varName += string(charData)
		}
	}
	for key, val := range dict {
		if strings.Index(varName, key) != -1 {
			varName = strings.Replace(varName, key, val, 1)
			// if expected value was found, clean the buffer and store replaced
			// value as CharData token
			buffer.Clean()
			return encoder.EncodeToken(xml.CharData(varName))
		}
	}
	// if expected value can't be found in a dictionary, just write
	// all nodes to XLS file and clean the buffer
	return buffer.Flush(encoder)
}

// WriteTo puts ZIP content to given writer (like a file of HTTP response)
func (doc *Docx) WriteTo(w io.Writer) (int64, error) {
	if doc.err != nil {
		return 0, doc.err
	}
	var total int64
	// store data in newly created zip file
	zipOut := zip.NewWriter(w)
	defer zipOut.Close()
	// we will look for document.xml file
	foundDoc := false
	// read data from a zip file
	for _, zipFile := range doc.zipReader.File {
		// create file inside zip archive
		w, err := zipOut.Create(zipFile.Name)
		if err != nil {
			return total, err
		}
		// read a file from inside of a zip archive
		r, err := zipFile.Open()
		if err != nil {
			return total, err
		}
		defer r.Close()

		// look for document.xml file, otherwise, just copy data
		if zipFile.Name != documentXML {
			n, err := io.Copy(w, r)
			total += n
			if err != nil {
				return total, err
			}
			continue
		}
		foundDoc = true
		decoder := xml.NewDecoder(r)
		encoder := xml.NewEncoder(w)
		buffer := make(Buffer, 0, 50)
		for {
			if err != nil {
				return total, err
			}
			// flush the buffer if we didn't find matching bracket in 50 tokens
			if cap(buffer)-len(buffer) == 0 {
				buffer.Flush(encoder)
			}
			token, err := decoder.RawToken()
			if err != nil {
				if err == io.EOF {
					break
				}
				return total, err
			}
			charData, isCharData := token.(xml.CharData)
			// we can look for brackets now even if it's not CharData token
			openingBracketIdx := bytes.IndexRune(charData, doc.openingBracket)
			closingBracketIdx := bytes.IndexRune(charData, doc.closingBracket)
			if len(buffer) == 0 {
				if !isCharData {
					err = encoder.EncodeToken(fixNS(token))
					if err != nil {
						return total, err
					}
					continue
				}
				if openingBracketIdx != -1 {
					buffer = append(buffer, xml.CopyToken(token))
				} else {
					err = encoder.EncodeToken(fixNS(token))
					if err != nil {
						return total, err
					}
				}
				if closingBracketIdx > openingBracketIdx {
					buffer.Process(encoder, doc.dict)
				}
			} else {
				buffer = append(buffer, xml.CopyToken(token))
				if !isCharData {
					continue
				}
				if closingBracketIdx != -1 { // TODO: this logic is broken
					buffer.Process(encoder, doc.dict)
				}
				// if closingBracketIdx < openingBracketIdx {
				// 	buffer.Process(encoder)
				// }
			}
		}
		err = buffer.Flush(encoder)
		if err != nil {
			return total, err
		}
		err = encoder.Flush()
		if err != nil {
			return total, err
		}
	}
	if !foundDoc {
		return total, fmt.Errorf("Invalid DOCX document: %s not found in the archive", documentXML)
	}
	return total, nil
}

// isWT checks if current token is <w:t> XML element
func isWT(name xml.Name) bool {
	return name.Space == "w" && name.Local == "t"
}

// fixNS is needed to write data in the same format as it was read
// Without it we would end up with broken DOCX
// not sure why `encoding/xml` stores data differently from how it was read
func fixNS(token xml.Token) xml.Token {
	switch token.(type) {
	case xml.StartElement:
		t := token.(xml.StartElement)
		t.Name = fixName(t.Name)
		for i := range t.Attr {
			t.Attr[i].Name = fixName(t.Attr[i].Name)
		}
		return t
	case xml.EndElement:
		t := token.(xml.EndElement)
		t.Name = fixName(t.Name)
		return t
	default:
		return token
	}
}

func fixName(name xml.Name) xml.Name {
	name.Local = name.Space + ":" + name.Local
	name.Space = ""
	return name
}
