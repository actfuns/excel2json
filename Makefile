BINARY := excel2json
CMD    := ./cmd/excel2json

.PHONY: build test race cover vet fmt lint clean

build:
	go build -o $(BINARY) $(CMD)

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
