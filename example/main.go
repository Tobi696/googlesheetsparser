package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"os"

	googlesheetsparser "github.com/Tobi696/google-sheets-parser"
	"golang.org/x/oauth2/jwt"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

type User struct {
	ID       uint
	Username string
	Name     string
	Email    string
	Password string
	Locale   string
}

type jwtConfig struct {
	Email        string   `json:"client_email"`
	PrivateKey   string   `json:"private_key"`
	PrivateKeyID string   `json:"private_key_id"`
	TokenURI     string   `json:"token_uri"`
	Scopes       []string `json:"scopes"`
}

func main() {
	var fileConf jwtConfig
	confFile, err := os.Open("credentials.json")
	if err != nil {
		log.Fatalf("Unable to read credentials file: %v", err)
	}
	defer confFile.Close()
	if err := json.NewDecoder(confFile).Decode(&fileConf); err != nil {
		log.Fatalf("Unable to parse credentials file: %v", err)
	}

	conf := &jwt.Config{
		Email:        fileConf.Email,
		PrivateKey:   []byte(fileConf.PrivateKey),
		PrivateKeyID: fileConf.PrivateKeyID,
		TokenURL:     fileConf.TokenURI,
		Scopes: []string{
			"https://www.googleapis.com/auth/spreadsheets.readonly",
		},
	}

	ctx := context.Background()
	srv, err := sheets.NewService(ctx, option.WithHTTPClient(conf.Client(ctx)))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	googlesheetsparser.Service = srv

	users, err := googlesheetsparser.ParsePageIntoStructSlice[User]("15PTbwnLdGJXb4kgLVVBtZ7HbK3QEj-olOxsY7XTzvCc", "Users")
	fmt.Println(users, err)
}
