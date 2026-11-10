BINARY  := excel2json
CMD     := ./cmd/excel2json
VERSION ?= $(shell git describe --tags --always --dirty 2>/dev/null || echo dev)
LDFLAGS := -X main.version=$(VERSION)

.PHONY: build test race cover vet fmt lint clean

build:
	go build -trimpath -ldflags "$(LDFLAGS)" -o $(BINARY) $(CMD)

test:
	go test ./...

race:
	go test -race ./...

cover:
	go test -cover ./...

vet:
	go vet ./...

fmt:
	go fmt ./...

lint:
	golangci-lint run

clean:
	rm -f $(BINARY)
