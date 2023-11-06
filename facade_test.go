package dgexcel

import (
	"encoding/json"
	dgctx "github.com/darwinOrg/go-common/context"
	dglogger "github.com/darwinOrg/go-logger"
	"testing"
)

type User struct {
	Name        string `excel:"name(名称);unique(true)"`
	Status      int    `excel:"name(状态);mapping(无效:0,有效:1)"`
	CreatedDate string `excel:"name(创建日期);date(01-02-06,2006-01-02)"`
}

func TestSimpleBindExcel2Struct(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")
	usersBytes, _ := json.Marshal(users)
	dglogger.Infof(ctx, "%s", string(usersBytes))
}

func TestBindExcelOfPrototype(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := BindExcelUsingTargetBuilder(ctx, "./users.xlsx", 1, 2, func() any {
		return &User{}
	})
	usersBytes, _ := json.Marshal(users)
	dglogger.Infof(ctx, "%s", string(usersBytes))
}

func TestSimpleExportStruct2XlsxFile(t *testing.T) {
	ctx := &dgctx.DgContext{TraceId: "123"}
	users, _ := SimpleBindExcel2Struct[User](ctx, "./users.xlsx")
	err := ExportStruct2XlsxFile(ctx, users, "./exported_users.xlsx")
	if err != nil {
		return
	}
}
