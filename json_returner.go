package main

import (
	"encoding/json"
	"log"
	"net/http"
)

func main() {
	http.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		// This is the JSON data we'll parse
		jsonData := []byte(`{
			"sheet1": {
				"headers": [
					"header A",
					"header B",
					"header C"
				],
				"data": {
					"column A": [
						"row 1": "value 1",
						"row 2": "value 2",
						"row 3": "value 3"
					],
					"column B": [
						"row 1": "value 4",
						"row 2": "value 5",
						"row 3": "value 6"
					],
				}
			}
		}`)

		// This is the variable where we'll store the parsed JSON
		var data map[string]interface{}

		// Parse the JSON
		err := json.Unmarshal(jsonData, &data)
		if err != nil {
			log.Fatal(err)
		}

		// Marshal the data back into JSON to return it
		jsonData, err = json.Marshal(data)
		if err != nil {
			log.Fatal(err)
		}

		w.Header().Set("Content-Type", "application/json")
		w.Write(jsonData)
	})

	log.Fatal(http.ListenAndServe(":8080", nil))
}