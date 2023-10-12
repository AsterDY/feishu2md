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

package core_test

import (
	"context"
	"fmt"
	"os"
	"testing"

	"github.com/Wsine/feishu2md/core"
)

func TestConvFromUrl(t *testing.T) {
	appid, apist := getIdAndSecretFromEnv()
	conf := &core.Config{
		Feishu: core.FeishuConfig{
			AppId:     appid,
			AppSecret: apist,
		},
		Output: core.OutputConfig{
			ImageDir:        ".test",
			TitleAsFilename: false,
			UseHTMLTags:     false,
			SkipImgDownload: false,
		},
	}
	url := os.Getenv("TEST_DOC")
	oclient := core.NewOClient(conf.Feishu.AppId, conf.Feishu.AppSecret)
	rets, err := core.CovertFromUrl(context.Background(), oclient, conf, url)
	if err != nil {
		t.Fatal(err)
	}
	fmt.Printf("%#v", rets)
}
