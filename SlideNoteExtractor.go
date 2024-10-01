package main

import (
	"archive/zip"
	"encoding/xml"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
)

// PowerPointのNotesのテキストを格納する構造体
type NotesSlide struct {
	XMLName xml.Name `xml:"notes"`
	CSld    CSld     `xml:"cSld"`
}

type CSld struct {
	SpTree SpTree `xml:"spTree"`
}

type SpTree struct {
	Shapes []Shape `xml:"sp"`
}

type Shape struct {
	TxBody TxBody `xml:"txBody"`
}

type TxBody struct {
	Paragraphs []Paragraph `xml:"p"`
}

type Paragraph struct {
	TextRuns []TextRun `xml:"r"`
}

type TextRun struct {
	Text string `xml:"t"`
}

func main() {
	// コマンドラインオプションの定義
	inputFile := flag.String("input", "", "入力PowerPointファイルのパス (.pptx)")
	outputFile := flag.String("output", "", "出力テキストファイルのパス")
	flag.Parse()

	// 入力ファイルと出力ファイルの確認
	if *inputFile == "" || *outputFile == "" {
		log.Fatalf("入力ファイルおよび出力ファイルを指定してください")
	}

	// ZIPファイル (pptx) を開く
	r, err := zip.OpenReader(*inputFile)
	if err != nil {
		log.Fatalf("PowerPointファイルの読み込みに失敗しました: %v", err)
	}
	defer r.Close()

	// 出力ファイルを作成
	out, err := os.Create(*outputFile)
	if err != nil {
		log.Fatalf("出力ファイルの作成に失敗しました: %v", err)
	}
	defer out.Close()

	// ZIP内のファイルを探索
	for _, f := range r.File {
		// ノートスライド (notesSlideN.xml) のファイルを探す
		if strings.HasPrefix(f.Name, "ppt/notesSlides/notesSlide") && strings.HasSuffix(f.Name, ".xml") {
			rc, err := f.Open()
			if err != nil {
				log.Fatalf("ファイルを開けませんでした: %v", err)
			}
			defer rc.Close()

			// XMLをパース
			notesSlide, err := extractNotesFromXML(rc)
			if err != nil {
				log.Printf("ノートの抽出に失敗しました: %v", err)
				continue
			}

			// 抽出したノートを出力ファイルに書き込み
			fmt.Fprintf(out, "ノートスライド %s:\n", filepath.Base(f.Name))
			for _, shape := range notesSlide.CSld.SpTree.Shapes {
				for _, p := range shape.TxBody.Paragraphs {
					var paragraphText string
					// 同じParagraph内のテキストを結合
					for _, r := range p.TextRuns {
						paragraphText += r.Text
					}
					if paragraphText != "" {
						fmt.Fprintf(out, "%s\n", paragraphText)
					}
				}
			}
			fmt.Fprintln(out, "")
		}
	}

	fmt.Println("ノートの抽出が完了しました")
}

// XMLからNotesのテキストを抽出
func extractNotesFromXML(reader io.Reader) (*NotesSlide, error) {
	decoder := xml.NewDecoder(reader)
	var notesSlide NotesSlide

	// XMLのパース処理
	for {
		tok, err := decoder.Token()
		if err == io.EOF {
			break
		} else if err != nil {
			return nil, err
		}

		// XMLトークンを確認しながら処理
		switch elem := tok.(type) {
		case xml.StartElement:
			// ノートスライドにマッチした場合、構造体にデコード
			if elem.Name.Local == "notes" {
				err = decoder.DecodeElement(&notesSlide, &elem)
				if err != nil {
					return nil, err
				}
			}
		}
	}

	return &notesSlide, nil
}
