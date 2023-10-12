// Copyright 2023 CloudWeGo Authors
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package core

import (
	"context"
	"fmt"
	"io"
	"regexp"
	"strings"

	"github.com/88250/lute"
	lark "github.com/larksuite/oapi-sdk-go/v3"
	larkdocx "github.com/larksuite/oapi-sdk-go/v3/service/docx/v1"
	larkdrive "github.com/larksuite/oapi-sdk-go/v3/service/drive/v1"
	larkim "github.com/larksuite/oapi-sdk-go/v3/service/im/v1"
	larkwiki "github.com/larksuite/oapi-sdk-go/v3/service/wiki/v2"
	"github.com/pkg/errors"
)

type ConvResult struct {
	Name       string
	FmtContent string
	RawContent string
	Images     []Image
}

type Image struct {
	Path    string
	Content io.Reader
	Key     string
}

func CovertFromUrl(ctx context.Context, oclient *lark.Client, config *Config, url string) (*ConvResult, error) {
	// configPath, err := GetConfigFilePath()
	// utils.CheckErr(err)
	// config, err := ReadConfigFromFile(configPath)
	// utils.CheckErr(err)

	// validate url
	reg := regexp.MustCompile("^https://[a-zA-Z0-9-.]+.(feishu.cn|larksuite.com)/(docx|wiki)/([a-zA-Z0-9]+)")
	matchResult := reg.FindStringSubmatch(url)
	if matchResult == nil || len(matchResult) != 4 {
		return nil, errors.Errorf("Invalid feishu/larksuite URL format")
	}

	// get doctype and token
	docType := matchResult[2]
	docToken := matchResult[3]
	fmt.Println("Captured document token:", docToken)

	ctx = context.WithValue(ctx, "output", config.Output)

	// for a wiki page, we need to renew docType and docToken first
	if docType == "wiki" {
		req := larkwiki.NewGetNodeSpaceReqBuilder().Token(docToken).ObjType(docType).Build()
		resp, err := oclient.Wiki.Space.GetNode(ctx, req)
		if err != nil {
			return nil, err
		}
		docType = *resp.Data.Node.ObjType
		docToken = *resp.Data.Node.ObjToken
	}
	if docType != "docx" {
		return nil, errors.New("convertor only support docx now!")
	}

	// get doc meta
	reqDoc := larkdocx.NewGetDocumentReqBuilder().DocumentId(docToken).Build()
	respDox, err := oclient.Docx.Document.Get(ctx, reqDoc)
	if err != nil {
		return nil, err
	}
	docx := respDox.Data.Document

	// get all blocks of the doc
	var blocks []*larkdocx.Block
	req := larkdocx.NewListDocumentBlockReqBuilder().DocumentId(docToken).Limit(-1).Build()
	iter, err := oclient.Docx.DocumentBlock.ListByIterator(ctx, req)
	if err != nil {
		return nil, err
	}
	for hasnext := true; hasnext; {
		ok, block, err := iter.Next()
		if err != nil {
			return nil, err
		}
		hasnext = ok
		if ok {
			blocks = append(blocks, block)
		}
	}

	// convert to markdown syntax
	parser := NewParser(ctx)
	markdown := parser.ParseDocxContent(docx, blocks)

	// handle images
	raw := markdown
	var images []Image
	for _, imgToken := range parser.ImgTokens {
		// NOTICE: imageToken is a media token of cloud driver, need upload to exchange to image key
		req := larkdrive.NewDownloadMediaReqBuilder().FileToken(imgToken).Build()
		resp, err := oclient.Drive.Media.Download(ctx, req)
		if err != nil {
			return nil, err
		}

		// TODO: store image token anywhere to avoid repeated uploading...
		im := larkim.NewCreateImageReqBuilder().
			Body(larkim.NewCreateImageReqBodyBuilder().
				ImageType(larkim.ImageTypeMessage).
				Image(resp.File).
				Build()).
			Build()
		resp2, err := oclient.Im.Image.Create(ctx, im)
		if err != nil {
			return nil, err
		}
		if !resp2.Success() {
			return nil, errors.New("upload image failed")
		}

		var image Image
		image.Key = *resp2.Data.ImageKey
		if config.Output.SkipImgDownload {
			markdown = strings.Replace(markdown, imgToken, image.Key, 1)
		} else {
			markdown = strings.Replace(markdown, imgToken, image.Path, 1)
		}
		images = append(images, image)
	}

	// formate markdown
	engine := lute.New(func(l *lute.Lute) {
		l.RenderOptions.AutoSpace = true
	})
	result := engine.FormatStr("md", markdown)

	title := *docx.Title
	mdName := fmt.Sprintf("%s.md", docToken)
	if config.Output.TitleAsFilename {
		mdName = fmt.Sprintf("%s.md", title)
	}

	// if err = os.WriteFile(mdName, ret, 0o644); err != nil {
	// 	return nil, err
	// }
	// fmt.Printf("Downloaded markdown file to %s\n", mdName)

	return &ConvResult{
		Name:       mdName,
		FmtContent: result,
		RawContent: raw,
		Images:     images,
	}, nil
}
