# SlideNoteExtractor
A program to extract Notes from a pptx file to a text file.

## 実行方法
### -output を指定しない場合

* この場合、出力ファイルは example.txt になります。

```
go run main.go -input="example.pptx"
```

### -output を指定した場合
* 指定されたファイルに出力されます。

```
go run main.go -input="example.pptx" -output="custom_output.txt"
```