package dgexcel

import (
	"fmt"
	dgcoll "github.com/darwinOrg/go-common/collection"
	dgctx "github.com/darwinOrg/go-common/context"
	dgerr "github.com/darwinOrg/go-common/enums/error"
	dglogger "github.com/darwinOrg/go-logger"
	"os"
	"strings"
)

func SimpleBindExcel2Struct[T any](ctx *dgctx.DgContext, filePath string) ([]*T, error) {
	return BindExcel2Struct[T](ctx, filePath, 1, 2)
}

func BindExcelUsingTargetBuilder(ctx *dgctx.DgContext, filePath string, headerRow int, dataStartRow int, targetBuilderFn func() any) ([]any, error) {
	file, err := os.Open(filePath)
	if err != nil {
		dglogger.Errorf(ctx, "open excel file error: %v", err)
		return nil, err
	}
	defer func(f *os.File) {
		err := f.Close()
		if err != nil {
			dglogger.Errorf(ctx, "close excel file error: %v", err)
		}
	}(file)

	parser, err := NewParser(targetBuilderFn())
	if err != nil {
		dglogger.Errorf(ctx, "new parser error: %v", err)
		return nil, err
	}

	rt, err := parser.ParseContent(file, headerRow, dataStartRow)
	if err != nil {
		dglogger.Errorf(ctx, "parse content error: %v", err)
		return nil, err
	}
	errList, has := rt.HasError()
	if has {
		var errs []string
		for k, v := range errList {
			dglogger.Warn(ctx, k, v)
			egs := dgcoll.MapToList(v, func(s string) string { return fmt.Sprintf("第%d行：%s", k, s) })
			errs = append(errs, egs...)
		}
		return nil, dgerr.SimpleDgError(strings.Join(errs, "；"))
	}

	ts, err := rt.FormatBaseTargetBuilder(targetBuilderFn)
	if err != nil {
		dglogger.Errorf(ctx, "bind to struct error: %v", err)
		return nil, err
	}

	return ts, nil
}

func BindExcel2Struct[T any](ctx *dgctx.DgContext, filePath string, headerRow int, dataStartRow int) ([]*T, error) {
	file, err := os.Open(filePath)
	if err != nil {
		dglogger.Errorf(ctx, "open excel file error: %v", err)
		return nil, err
	}
	defer func(f *os.File) {
		err := f.Close()
		if err != nil {
			dglogger.Errorf(ctx, "close excel file error: %v", err)
		}
	}(file)

	t := new(T)
	processor, err := NewParser(t)
	if err != nil {
		dglogger.Errorf(ctx, "new parser error: %v", err)
		return nil, err
	}

	rt, err := processor.ParseContent(file, headerRow, dataStartRow)
	if err != nil {
		dglogger.Errorf(ctx, "parse content error: %v", err)
		return nil, err
	}
	errList, has := rt.HasError()
	if has {
		var errs []string
		for k, v := range errList {
			dglogger.Warn(ctx, k, v)
			egs := dgcoll.MapToList(v, func(s string) string { return fmt.Sprintf("第%d行：%s", k, s) })
			errs = append(errs, egs...)
		}
		return nil, dgerr.SimpleDgError(strings.Join(errs, "；"))
	}

	ts := make([]*T, 0)
	err = rt.Format(&ts)
	if err != nil {
		dglogger.Errorf(ctx, "bind to struct error: %v", err)
		return nil, err
	}

	return ts, nil
}

func ExportStruct2XlsxFile(ctx *dgctx.DgContext, v any, filePath string) error {
	xlsx, err := ExportStruct2Xlsx(v)
	if err != nil {
		dglogger.Errorf(ctx, "export struct to xlsx file error: %v", err)
		return err
	}

	err = xlsx.SaveAs(filePath)
	if err != nil {
		dglogger.Errorf(ctx, "save exported xlsx file error: %v", err)
		return err
	}

	return nil
}
