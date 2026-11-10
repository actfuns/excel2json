package main

import (
	"runtime/debug"

	"github.com/actfuns/excel2json/internal/app"
)

var version = "dev" // overridden with -ldflags "-X main.version=..."

func main() {
	if version == "dev" {
		if info, ok := debug.ReadBuildInfo(); ok && info.Main.Version != "" && info.Main.Version != "(devel)" {
			version = info.Main.Version
		}
	}
	app.Main(version)
}
