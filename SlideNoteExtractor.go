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
type NotesText struct {
	XMLName    xml.Name    `xml:"txBody"`
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
			fmt.Printf("ノートスライドを処理しています: %s\n", f.Name)
			rc, err := f.Open()
			if err != nil {
				log.Fatalf("ファイルを開けませんでした: %v", err)
			}
			defer rc.Close()

			// XMLをパース
			notesText, err := extractNotesFromXML(rc)
			if err != nil {
				log.Printf("ノートの抽出に失敗しました: %v", err)
				continue
			}

			// 抽出したノートを出力ファイルに書き込み
			fmt.Fprintf(out, "ノートスライド %s:\n", filepath.Base(f.Name))
			for _, p := range notesText.Paragraphs {
				for _, r := range p.TextRuns {
					if r.Text != "" {
						fmt.Fprintf(out, "%s\n", r.Text)
					}
				}
			}
			fmt.Fprintln(out, "")
		}
	}

	fmt.Println("ノートの抽出が完了しました")
}

// XMLからNotesのテキストを抽出
func extractNotesFromXML(reader io.Reader) (*NotesText, error) {
	decoder := xml.NewDecoder(reader)
	var notes NotesText
	err := decoder.Decode(&notes)
	if err != nil {
		return nil, err
	}
	return &notes, nil
}
