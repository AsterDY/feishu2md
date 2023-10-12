package core

import (
	"context"
	"fmt"
	"strings"

	"github.com/Wsine/feishu2md/utils"
	lark "github.com/larksuite/oapi-sdk-go/v3/service/docx/v1"
	"github.com/olekukonko/tablewriter"
)

type Parser struct {
	ctx       context.Context
	ImgTokens []string
	blockMap  map[string]*lark.Block
}

func NewParser(ctx context.Context) *Parser {
	return &Parser{
		ctx:       ctx,
		ImgTokens: make([]string, 0),
		blockMap:  make(map[string]*lark.Block),
	}
}

type DocxBlockType int

const (
	DocxBlockTypePage           DocxBlockType = 1   // 文档 Block
	DocxBlockTypeText           DocxBlockType = 2   // 文本 Block
	DocxBlockTypeHeading1       DocxBlockType = 3   // 一级标题 Block
	DocxBlockTypeHeading2       DocxBlockType = 4   // 二级标题 Block
	DocxBlockTypeHeading3       DocxBlockType = 5   // 三级标题 Block
	DocxBlockTypeHeading4       DocxBlockType = 6   // 四级标题 Block
	DocxBlockTypeHeading5       DocxBlockType = 7   // 五级标题 Block
	DocxBlockTypeHeading6       DocxBlockType = 8   // 六级标题 Block
	DocxBlockTypeHeading7       DocxBlockType = 9   // 七级标题 Block
	DocxBlockTypeHeading8       DocxBlockType = 10  // 八级标题 Block
	DocxBlockTypeHeading9       DocxBlockType = 11  // 九级标题 Block
	DocxBlockTypeBullet         DocxBlockType = 12  // 无序列表 Block
	DocxBlockTypeOrdered        DocxBlockType = 13  // 有序列表 Block
	DocxBlockTypeCode           DocxBlockType = 14  // 代码块 Block
	DocxBlockTypeQuote          DocxBlockType = 15  // 引用 Block
	DocxBlockTypeEquation       DocxBlockType = 16  // 公式 Block
	DocxBlockTypeTodo           DocxBlockType = 17  // 任务 Block
	DocxBlockTypeBitable        DocxBlockType = 18  // 多维表格 Block
	DocxBlockTypeCallout        DocxBlockType = 19  // 高亮块 Block
	DocxBlockTypeChatCard       DocxBlockType = 20  // 群聊卡片 Block
	DocxBlockTypeDiagram        DocxBlockType = 21  // 流程图/UML Block
	DocxBlockTypeDivider        DocxBlockType = 22  // 分割线 Block
	DocxBlockTypeFile           DocxBlockType = 23  // 文件 Block
	DocxBlockTypeGrid           DocxBlockType = 24  // 分栏 Block
	DocxBlockTypeGridColumn     DocxBlockType = 25  // 分栏列 Block
	DocxBlockTypeIframe         DocxBlockType = 26  // 内嵌 Block
	DocxBlockTypeImage          DocxBlockType = 27  // 图片 Block
	DocxBlockTypeISV            DocxBlockType = 28  // 三方 Block
	DocxBlockTypeMindnote       DocxBlockType = 29  // 思维笔记 Block
	DocxBlockTypeSheet          DocxBlockType = 30  // 电子表格 Block
	DocxBlockTypeTable          DocxBlockType = 31  // 表格 Block
	DocxBlockTypeTableCell      DocxBlockType = 32  // 单元格 Block
	DocxBlockTypeView           DocxBlockType = 33  // 视图 Block
	DocxBlockTypeQuoteContainer DocxBlockType = 34  // 引用容器 Block
	DocxBlockTypeTask           DocxBlockType = 35  // 任务容器 Block
	DocxBlockTypeOKR            DocxBlockType = 36  // OKR容器 Block
	DocxBlockTypeOKRObjective   DocxBlockType = 37  // OKR Objective容器 Block
	DocxBlockTypeOKRKeyResult   DocxBlockType = 38  // OKR KeyResult容器 Block
	DocxBlockTypeProgress       DocxBlockType = 39  // Progress容器 Block
	DocxBlockTypeUndefined      DocxBlockType = 999 // 未支持 Block
)

type DocxCodeLanguage int64

const (
	DocxCodeLanguagePlainText    DocxCodeLanguage = 1  // PlainText
	DocxCodeLanguageABAP         DocxCodeLanguage = 2  // ABAP
	DocxCodeLanguageAda          DocxCodeLanguage = 3  // Ada
	DocxCodeLanguageApache       DocxCodeLanguage = 4  // Apache
	DocxCodeLanguageApex         DocxCodeLanguage = 5  // Apex
	DocxCodeLanguageAssembly     DocxCodeLanguage = 6  // Assembly
	DocxCodeLanguageBash         DocxCodeLanguage = 7  // Bash
	DocxCodeLanguageCSharp       DocxCodeLanguage = 8  // CSharp
	DocxCodeLanguageCPlusPlus    DocxCodeLanguage = 9  // C++
	DocxCodeLanguageC            DocxCodeLanguage = 10 // C
	DocxCodeLanguageCOBOL        DocxCodeLanguage = 11 // COBOL
	DocxCodeLanguageCSS          DocxCodeLanguage = 12 // CSS
	DocxCodeLanguageCoffeeScript DocxCodeLanguage = 13 // CoffeeScript
	DocxCodeLanguageD            DocxCodeLanguage = 14 // D
	DocxCodeLanguageDart         DocxCodeLanguage = 15 // Dart
	DocxCodeLanguageDelphi       DocxCodeLanguage = 16 // Delphi
	DocxCodeLanguageDjango       DocxCodeLanguage = 17 // Django
	DocxCodeLanguageDockerfile   DocxCodeLanguage = 18 // Dockerfile
	DocxCodeLanguageErlang       DocxCodeLanguage = 19 // Erlang
	DocxCodeLanguageFortran      DocxCodeLanguage = 20 // Fortran
	DocxCodeLanguageFoxPro       DocxCodeLanguage = 21 // FoxPro
	DocxCodeLanguageGo           DocxCodeLanguage = 22 // Go
	DocxCodeLanguageGroovy       DocxCodeLanguage = 23 // Groovy
	DocxCodeLanguageHTML         DocxCodeLanguage = 24 // HTML
	DocxCodeLanguageHTMLBars     DocxCodeLanguage = 25 // HTMLBars
	DocxCodeLanguageHTTP         DocxCodeLanguage = 26 // HTTP
	DocxCodeLanguageHaskell      DocxCodeLanguage = 27 // Haskell
	DocxCodeLanguageJSON         DocxCodeLanguage = 28 // JSON
	DocxCodeLanguageJava         DocxCodeLanguage = 29 // Java
	DocxCodeLanguageJavaScript   DocxCodeLanguage = 30 // JavaScript
	DocxCodeLanguageJulia        DocxCodeLanguage = 31 // Julia
	DocxCodeLanguageKotlin       DocxCodeLanguage = 32 // Kotlin
	DocxCodeLanguageLateX        DocxCodeLanguage = 33 // LateX
	DocxCodeLanguageLisp         DocxCodeLanguage = 34 // Lisp
	DocxCodeLanguageLogo         DocxCodeLanguage = 35 // Logo
	DocxCodeLanguageLua          DocxCodeLanguage = 36 // Lua
	DocxCodeLanguageMATLAB       DocxCodeLanguage = 37 // MATLAB
	DocxCodeLanguageMakefile     DocxCodeLanguage = 38 // Makefile
	DocxCodeLanguageMarkdown     DocxCodeLanguage = 39 // Markdown
	DocxCodeLanguageNginx        DocxCodeLanguage = 40 // Nginx
	DocxCodeLanguageObjective    DocxCodeLanguage = 41 // Objective
	DocxCodeLanguageOpenEdgeABL  DocxCodeLanguage = 42 // OpenEdgeABL
	DocxCodeLanguagePHP          DocxCodeLanguage = 43 // PHP
	DocxCodeLanguagePerl         DocxCodeLanguage = 44 // Perl
	DocxCodeLanguagePostScript   DocxCodeLanguage = 45 // PostScript
	DocxCodeLanguagePower        DocxCodeLanguage = 46 // Power
	DocxCodeLanguageProlog       DocxCodeLanguage = 47 // Prolog
	DocxCodeLanguageProtoBuf     DocxCodeLanguage = 48 // ProtoBuf
	DocxCodeLanguagePython       DocxCodeLanguage = 49 // Python
	DocxCodeLanguageR            DocxCodeLanguage = 50 // R
	DocxCodeLanguageRPG          DocxCodeLanguage = 51 // RPG
	DocxCodeLanguageRuby         DocxCodeLanguage = 52 // Ruby
	DocxCodeLanguageRust         DocxCodeLanguage = 53 // Rust
	DocxCodeLanguageSAS          DocxCodeLanguage = 54 // SAS
	DocxCodeLanguageSCSS         DocxCodeLanguage = 55 // SCSS
	DocxCodeLanguageSQL          DocxCodeLanguage = 56 // SQL
	DocxCodeLanguageScala        DocxCodeLanguage = 57 // Scala
	DocxCodeLanguageScheme       DocxCodeLanguage = 58 // Scheme
	DocxCodeLanguageScratch      DocxCodeLanguage = 59 // Scratch
	DocxCodeLanguageShell        DocxCodeLanguage = 60 // Shell
	DocxCodeLanguageSwift        DocxCodeLanguage = 61 // Swift
	DocxCodeLanguageThrift       DocxCodeLanguage = 62 // Thrift
	DocxCodeLanguageTypeScript   DocxCodeLanguage = 63 // TypeScript
	DocxCodeLanguageVBScript     DocxCodeLanguage = 64 // VBScript
	DocxCodeLanguageVisual       DocxCodeLanguage = 65 // Visual
	DocxCodeLanguageXML          DocxCodeLanguage = 66 // XML
	DocxCodeLanguageYAML         DocxCodeLanguage = 67 // YAML
)

var DocxCodeLang2MdStr = map[DocxCodeLanguage]string{
	DocxCodeLanguagePlainText:    "",
	DocxCodeLanguageABAP:         "abap",
	DocxCodeLanguageAda:          "ada",
	DocxCodeLanguageApache:       "apache",
	DocxCodeLanguageApex:         "apex",
	DocxCodeLanguageAssembly:     "assembly",
	DocxCodeLanguageBash:         "bash",
	DocxCodeLanguageCSharp:       "csharp",
	DocxCodeLanguageCPlusPlus:    "cpp",
	DocxCodeLanguageC:            "c",
	DocxCodeLanguageCOBOL:        "cobol",
	DocxCodeLanguageCSS:          "css",
	DocxCodeLanguageCoffeeScript: "coffeescript",
	DocxCodeLanguageD:            "d",
	DocxCodeLanguageDart:         "dart",
	DocxCodeLanguageDelphi:       "delphi",
	DocxCodeLanguageDjango:       "django",
	DocxCodeLanguageDockerfile:   "dockerfile",
	DocxCodeLanguageErlang:       "erlang",
	DocxCodeLanguageFortran:      "fortran",
	DocxCodeLanguageFoxPro:       "foxpro",
	DocxCodeLanguageGo:           "go",
	DocxCodeLanguageGroovy:       "groovy",
	DocxCodeLanguageHTML:         "html",
	DocxCodeLanguageHTMLBars:     "htmlbars",
	DocxCodeLanguageHTTP:         "http",
	DocxCodeLanguageHaskell:      "haskell",
	DocxCodeLanguageJSON:         "json",
	DocxCodeLanguageJava:         "java",
	DocxCodeLanguageJavaScript:   "javascript",
	DocxCodeLanguageJulia:        "julia",
	DocxCodeLanguageKotlin:       "kotlin",
	DocxCodeLanguageLateX:        "latex",
	DocxCodeLanguageLisp:         "lisp",
	DocxCodeLanguageLogo:         "logo",
	DocxCodeLanguageLua:          "lua",
	DocxCodeLanguageMATLAB:       "matlab",
	DocxCodeLanguageMakefile:     "makefile",
	DocxCodeLanguageMarkdown:     "markdown",
	DocxCodeLanguageNginx:        "nginx",
	DocxCodeLanguageObjective:    "objectivec",
	DocxCodeLanguageOpenEdgeABL:  "openedge-abl",
	DocxCodeLanguagePHP:          "php",
	DocxCodeLanguagePerl:         "perl",
	DocxCodeLanguagePostScript:   "postscript",
	DocxCodeLanguagePower:        "powershell",
	DocxCodeLanguageProlog:       "prolog",
	DocxCodeLanguageProtoBuf:     "protobuf",
	DocxCodeLanguagePython:       "python",
	DocxCodeLanguageR:            "r",
	DocxCodeLanguageRPG:          "rpg",
	DocxCodeLanguageRuby:         "ruby",
	DocxCodeLanguageRust:         "rust",
	DocxCodeLanguageSAS:          "sas",
	DocxCodeLanguageSCSS:         "scss",
	DocxCodeLanguageSQL:          "sql",
	DocxCodeLanguageScala:        "scala",
	DocxCodeLanguageScheme:       "scheme",
	DocxCodeLanguageScratch:      "scratch",
	DocxCodeLanguageShell:        "shell",
	DocxCodeLanguageSwift:        "swift",
	DocxCodeLanguageThrift:       "thrift",
	DocxCodeLanguageTypeScript:   "typescript",
	DocxCodeLanguageVBScript:     "vbscript",
	DocxCodeLanguageVisual:       "vbnet",
	DocxCodeLanguageXML:          "xml",
	DocxCodeLanguageYAML:         "yaml",
}

// =============================================================
// Parser utils
// =============================================================

func renderMarkdownTable(data [][]string) string {
	builder := &strings.Builder{}
	table := tablewriter.NewWriter(builder)
	table.SetCenterSeparator("|")
	table.SetAutoWrapText(false)
	table.SetAutoFormatHeaders(false)
	table.SetAutoMergeCells(false)
	table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
	table.SetHeader(data[0])
	table.AppendBulk(data[1:])
	table.Render()
	return builder.String()
}

// =============================================================
// Parse the new version of document (docx)
// =============================================================

func (p *Parser) ParseDocxContent(doc *lark.Document, blocks []*lark.Block) string {
	for _, block := range blocks {
		p.blockMap[*block.BlockId] = block
	}

	entryBlock := p.blockMap[*doc.DocumentId]
	return p.ParseDocxBlock(entryBlock, 0)
}

func (p *Parser) ParseDocxBlock(b *lark.Block, indentLevel int) string {
	buf := new(strings.Builder)
	buf.WriteString(strings.Repeat("\t", indentLevel))
	switch DocxBlockType(*b.BlockType) {
	case DocxBlockTypePage:
		buf.WriteString(p.ParseDocxBlockPage(b))
	case DocxBlockTypeText:
		buf.WriteString(p.ParseDocxBlockText(b.Text))
	case DocxBlockTypeHeading1:
		buf.WriteString("# ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading1))
	case DocxBlockTypeHeading2:
		buf.WriteString("## ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading2))
	case DocxBlockTypeHeading3:
		buf.WriteString("### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading3))
	case DocxBlockTypeHeading4:
		buf.WriteString("#### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading4))
	case DocxBlockTypeHeading5:
		buf.WriteString("##### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading5))
	case DocxBlockTypeHeading6:
		buf.WriteString("###### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading6))
	case DocxBlockTypeHeading7:
		buf.WriteString("####### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading7))
	case DocxBlockTypeHeading8:
		buf.WriteString("######## ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading8))
	case DocxBlockTypeHeading9:
		buf.WriteString("######### ")
		buf.WriteString(p.ParseDocxBlockText(b.Heading9))
	case DocxBlockTypeBullet:
		buf.WriteString(p.ParseDocxBlockBullet(b, indentLevel))
	case DocxBlockTypeOrdered:
		buf.WriteString(p.ParseDocxBlockOrdered(b, indentLevel))
	case DocxBlockTypeCode:
		var style string
		if b.Code.Style.Language != nil {
			style = DocxCodeLang2MdStr[DocxCodeLanguage(*b.Code.Style.Language)]
		}
		buf.WriteString("```" + style + "\n")
		buf.WriteString(strings.TrimSpace(p.ParseDocxBlockText(b.Code)))
		buf.WriteString("\n```\n")
	case DocxBlockTypeQuote:
		buf.WriteString("> ")
		buf.WriteString(p.ParseDocxBlockText(b.Quote))
	case DocxBlockTypeEquation:
		buf.WriteString("$$\n")
		buf.WriteString(p.ParseDocxBlockText(b.Equation))
		buf.WriteString("\n$$\n")
	case DocxBlockTypeTodo:
		if s := b.Todo.Style.Done; s != nil && *s {
			buf.WriteString("- [x] ")
		} else {
			buf.WriteString("- [ ] ")
		}
		buf.WriteString(p.ParseDocxBlockText(b.Todo))
	case DocxBlockTypeDivider:
		buf.WriteString("---\n")
	case DocxBlockTypeImage:
		buf.WriteString(p.ParseDocxBlockImage(b.Image))
	case DocxBlockTypeTableCell:
		buf.WriteString(p.ParseDocxBlockTableCell(b))
	case DocxBlockTypeTable:
		buf.WriteString(p.ParseDocxBlockTable(b.Table))
	case DocxBlockTypeQuoteContainer:
		buf.WriteString(p.ParseDocxBlockQuoteContainer(b))
	default:
	}
	return buf.String()
}

func (p *Parser) ParseDocxBlockPage(b *lark.Block) string {
	buf := new(strings.Builder)

	buf.WriteString("# ")
	buf.WriteString(p.ParseDocxBlockText(b.Page))
	buf.WriteString("\n")

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, 0))
		buf.WriteString("\n")
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockText(b *lark.Text) string {
	buf := new(strings.Builder)
	numElem := len(b.Elements)
	for _, e := range b.Elements {
		inline := numElem > 1
		buf.WriteString(p.ParseDocxTextElement(e, inline))
	}
	buf.WriteString("\n")
	return buf.String()
}

func (p *Parser) ParseDocxTextElement(e *lark.TextElement, inline bool) string {
	buf := new(strings.Builder)
	if e.TextRun != nil {
		buf.WriteString(p.ParseDocxTextElementTextRun(e.TextRun))
	}
	if e.MentionUser != nil {
		buf.WriteString(*e.MentionUser.UserId)
	}
	if e.MentionDoc != nil {
		var title string
		if e.MentionDoc.Title != nil {
			title = *e.MentionDoc.Title
		}
		var url string
		if e.MentionDoc.Url != nil {
			url = *e.MentionDoc.Url
		}
		buf.WriteString(fmt.Sprintf("[%s](%s)", title, utils.UnescapeURL(url)))
	}
	if e.Equation != nil {
		symbol := "$$"
		if inline {
			symbol = "$"
		}
		var cont string
		if e.Equation.Content != nil {
			cont = *e.Equation.Content
		}
		buf.WriteString(symbol + strings.TrimSuffix(cont, "\n") + symbol)
	}
	return buf.String()
}

func (p *Parser) ParseDocxTextElementTextRun(tr *lark.TextRun) string {
	buf := new(strings.Builder)
	postWrite := ""
	if style := tr.TextElementStyle; style != nil {
		useHTMLTags := NewConfig("", "").Output.UseHTMLTags
		if p.ctx.Value("output") != nil {
			useHTMLTags = p.ctx.Value("output").(OutputConfig).UseHTMLTags
		}
		if v := style.Bold; v != nil && *v {
			if useHTMLTags {
				buf.WriteString("<strong>")
				postWrite = "</strong>"
			} else {
				buf.WriteString("**")
				postWrite = "**"
			}
		} else if v := style.Italic; v != nil && *v {
			if useHTMLTags {
				buf.WriteString("<em>")
				postWrite = "</em>"
			} else {
				buf.WriteString("_")
				postWrite = "_"
			}
		} else if v := style.Strikethrough; v != nil && *v {
			if useHTMLTags {
				buf.WriteString("<del>")
				postWrite = "</del>"
			} else {
				buf.WriteString("~~")
				postWrite = "~~"
			}
		} else if v := style.Underline; v != nil && *v {
			buf.WriteString("<u>")
			postWrite = "</u>"
		} else if v := style.InlineCode; v != nil && *v {
			buf.WriteString("`")
			postWrite = "`"
		} else if link := style.Link; link != nil {
			buf.WriteString("[")
			var url string
			if link.Url != nil {
				url = *link.Url
			}
			postWrite = fmt.Sprintf("](%s)", utils.UnescapeURL(url))
		}
	}
	if tr.Content != nil {
		buf.WriteString(*tr.Content)
	}
	buf.WriteString(postWrite)
	return buf.String()
}

func (p *Parser) ParseDocxBlockImage(img *lark.Image) string {
	buf := new(strings.Builder)
	var url = *img.Token
	buf.WriteString(fmt.Sprintf("![](%s)", url))
	buf.WriteString("\n")
	p.ImgTokens = append(p.ImgTokens, *img.Token)
	return buf.String()
}

func (p *Parser) ParseDocxBlockBullet(b *lark.Block, indentLevel int) string {
	buf := new(strings.Builder)

	buf.WriteString("- ")
	buf.WriteString(p.ParseDocxBlockText(b.Bullet))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockOrdered(b *lark.Block, indentLevel int) string {
	buf := new(strings.Builder)

	// calculate order and indent level
	parent := p.blockMap[*b.ParentId]
	order := 1
	for idx, child := range parent.Children {
		if child == *b.BlockId {
			for i := idx - 1; i >= 0; i-- {
				if v := p.blockMap[parent.Children[i]].BlockType; v != nil && *v == int(DocxBlockTypeOrdered) {
					order += 1
				} else {
					break
				}
			}
			break
		}
	}

	buf.WriteString(fmt.Sprintf("%d. ", order))
	buf.WriteString(p.ParseDocxBlockText(b.Ordered))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTableCell(b *lark.Block) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		block := p.blockMap[child]
		content := p.ParseDocxBlock(block, 0)
		buf.WriteString(content)
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTable(t *lark.Table) string {
	// - First row as header
	// - Ignore cell merging
	s := t.Property.ColumnSize
	if s == nil || *s == 0 {
		return ""
	}
	var rows [][]string
	for i, blockId := range t.Cells {
		block := p.blockMap[blockId]
		cellContent := p.ParseDocxBlock(block, 0)
		cellContent = strings.ReplaceAll(cellContent, "\n", "")
		rowIndex := i / *s
		if len(rows) < int(rowIndex)+1 {
			rows = append(rows, []string{})
		}
		rows[rowIndex] = append(rows[rowIndex], cellContent)
	}

	buf := new(strings.Builder)
	buf.WriteString(renderMarkdownTable(rows))
	buf.WriteString("\n")
	return buf.String()
}

func (p *Parser) ParseDocxBlockQuoteContainer(b *lark.Block) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		block := p.blockMap[child]
		buf.WriteString("> ")
		buf.WriteString(p.ParseDocxBlock(block, 0))
	}

	return buf.String()
}
