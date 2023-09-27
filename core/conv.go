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
	"github.com/Wsine/feishu2md/utils"
	"github.com/pkg/errors"
)

type ConvResult struct {
	Name       string
	NewContent string
	RawContent string
	Images     []Image
}

type Image struct {
	Name    string
	Content io.Reader
}

func CovertFromUrl(config *Config, url string) (*ConvResult, error) {
	// configPath, err := GetConfigFilePath()
	// utils.CheckErr(err)
	// config, err := ReadConfigFromFile(configPath)
	// utils.CheckErr(err)

	reg := regexp.MustCompile("^https://[a-zA-Z0-9-.]+.(feishu.cn|larksuite.com)/(docx|wiki)/([a-zA-Z0-9]+)")
	matchResult := reg.FindStringSubmatch(url)
	if matchResult == nil || len(matchResult) != 4 {
		return nil, errors.Errorf("Invalid feishu/larksuite URL format")
	}

	domain := matchResult[1]
	docType := matchResult[2]
	docToken := matchResult[3]
	fmt.Println("Captured document token:", docToken)

	ctx := context.WithValue(context.Background(), "output", config.Output)

	client := NewClient(
		config.Feishu.AppId, config.Feishu.AppSecret, domain,
	)

	// for a wiki page, we need to renew docType and docToken first
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		utils.CheckErr(err)
		docType = node.ObjType
		docToken = node.ObjToken
	}

	docx, blocks, err := client.GetDocxContent(ctx, docToken)
	utils.CheckErr(err)

	parser := NewParser(ctx)

	markdown := parser.ParseDocxContent(docx, blocks)
	raw := markdown

	var images []Image
	if !config.Output.SkipImgDownload {
		for _, imgToken := range parser.ImgTokens {
			image, err := client.GetImage(ctx, imgToken, config.Output.ImageDir)
			if err != nil {
				return nil, err
			}
			//OPT: replace all in once
			markdown = strings.Replace(markdown, imgToken, image.Name, 1)
			images = append(images, image)
		}
	}

	engine := lute.New(func(l *lute.Lute) {
		l.RenderOptions.AutoSpace = true
	})
	result := engine.FormatStr("md", markdown)

	title := docx.Title
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
		NewContent: result,
		RawContent: raw,
		Images:     images,
	}, nil
}
