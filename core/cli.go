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
	"strings"
	"sync"

	lark "github.com/larksuite/oapi-sdk-go/v3"
)

var oClients sync.Map

func NewOClient(appID, appSecret string) *lark.Client {
	h := strings.Join([]string{appID, appSecret}, "-")
	if cli, ok := oClients.Load(h); ok {
		return cli.(*lark.Client)
	} else {
		cli := lark.NewClient(appID, appSecret)
		oClients.Store(h, cli)
		return cli
	}
}
