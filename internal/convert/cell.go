package convert

import (
	"fmt"

	"github.com/actfuns/excel2json/internal/schema"
	"github.com/actfuns/excel2json/internal/typing"
)

type cachedType struct {
	t   *typing.Type
	err error
}

func (e *Engine) parseTypeCached(colType string) (*typing.Type, error) {
	if v, ok := e.typeCache.Load(colType); ok {
		c := v.(cachedType)
		return c.t, c.err
	}
	t, err := typing.Parse(colType)
	e.typeCache.Store(colType, cachedType{t: t, err: err})
	return t, err
}

type cachedSpec struct {
	spec *schema.Spec
	err  error
}

func (e *Engine) parseSpec(annotation string) (*schema.Spec, error) {
	if v, ok := e.specCache.Load(annotation); ok {
		c := v.(cachedSpec)
		return c.spec, c.err
	}
	spec, err := schema.Parse(annotation, e.opts.TypeDefs)
	e.specCache.Store(annotation, cachedSpec{spec: spec, err: err})
	return spec, err
}

func (e *Engine) convertTyped(colType, cell string) (any, error) {
	t, err := e.parseTypeCached(colType)
	if err != nil {
		return nil, err
	}
	if cell == "" {
		if e.opts.RequireByDefault {
			return nil, fmt.Errorf("field cannot be empty (type %s)", t)
		}
		return typing.ZeroValue(t), nil
	}
	return e.conv.ConvertString(cell, t)
}
