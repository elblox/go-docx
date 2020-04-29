package docx

import (
	"archive/zip"
	"bytes"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"testing"
)

const fileName = "go-docx-test.docx"

var dict = map[string]string{
	"[simple]":                 "SiMPlE",
	"[with_color]":             "WiTh CoLoR",
	"[with_overlapping_color]": "WiTh OvErLaPiNg CoLoR",
}

func TestOpenAndSaveDocx(t *testing.T) {
	buf := new(bytes.Buffer)
	f, err := os.Open(fileName)
	if err != nil {
		t.Fatal(err)
	}
	fi, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	n, err := New(f, fi.Size()).Replace(dict).WriteTo(buf)
	if err != nil {
		log.Fatal(err)
	}
	if n == 0 {
		t.Error("Zero bytes written")
	}
	if http.DetectContentType(buf.Bytes()) != "application/zip" {
		t.Fatal("Unexpected content type")
	}

	// check zip content
	reader, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	for _, zipFile := range reader.File {
		if zipFile.Name != documentXML {
			continue
		}
		r, err := zipFile.Open()
		if err != nil {
			t.Fatal(err)
		}
		content, err := ioutil.ReadAll(r)
		if err != nil {
			t.Fatal(err)
		}
		for _, val := range dict {
			if !bytes.Contains(content, []byte(val)) {
				t.Errorf("Can't find value `%s` in %s", val, documentXML)
			}
		}
	}

	// write to file
	file, err := os.Create("_" + fileName)
	if err != nil {
		t.Fatal(err)
	}
	_, err = io.Copy(file, buf)
	if err != nil {
		t.Fatal(err)
	}
}
