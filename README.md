# docxtpl_ex
fast api docxtpl html, dynamic merge vertical and more

# cp
https://github.com/elapouya/python-docx-template
https://github.com/tsy19900929/docxtpl_dynamic_vertical_merging
https://github.com/dfop02/html4docx


# run
```bash
uvicorn main:app --reload --port 8000
```

# test
```bash
curl --request POST \
  --url http://127.0.0.1:8000/docx/ \
  --header 'content-type: application/json' \
  --data '{
  "file": "dnn.docx",
  "content": {
    "merge vertical hierarchy": {
      "token": "merge_vertical_hierarchy",
      "merges_child": [
        {
          "token": "m_nn",
          "atr": "childs",
          "column": "nn"
        },
        {
          "token": "m_c1",
          "atr": "childs",
          "column": "c1"
          
        }
        
      ],
      "data": [
        {
          "name": "X: 03",
          "childs": [
            {
              "name": "X1",
              "c1": "27/10/2024\n00:05:38",
              "c2": "28/10/2024\n08:35:25",
              "nn": "- abcdef\n- qwerty"
            },
            {
              "name": "X2",
              "c1": "27/10/2024\n00:05:38",
              "c2": "28/10/2024\n08:35:25",
              "nn": "- abcdef"
            },
            {
              "name": "X3",
              "c1": "28/10/2024\n00:05:38",
              "c2": "28/10/2024\n08:35:25",
              "nn": "- abcdef"
            }
          ]
        },
        {
          "name": "Z: 02",
          "childs": [
            {
              "name": "Z1",
              "c1": "27/10/2024\n00:05:38",
              "c2": "28/10/2024\n08:35:25",
              "nn": "- qwerty"
            },
            {
              "name": "Z2",
              "c1": "27/10/2024\n00:05:38",
              "c2": "28/10/2024\n08:35:25",
              "nn": "- qwerty"
            }
          ]
        }
      ]
    },
    "arr1": {
      "token": "arr1",
      "data": [
        {
          "name": "tung nguyen",
          "pos": "dev"
        },
        {
          "name": "diep nguyen",
          "pos": "tester"
        },
        {
          "name": "ha phuong nguyen",
          "pos": "ba",
          "mb": "123456"
        }
      ]
    },
    "html": {
      "token": "html",
      "html": true,
      "data": "<i>Xin chào</i> Viet Nam"
    },
    "html2": {
      "token": "html2",
      "html": true,
      "data": "<a href=\"https://github.com/BonsoirDiep/docxtpl_ex\">docxtpl_ex</a>"
    },
    "merge vertical 1": {
      "token": "merge_vertical_1",
      "merges": [
        {
          "token": "B"
        }
      ],
      "data": [
        1,
        2,
        2,
        3,
        3,
        3,
        4,
        4,
        4,
        4,
        5
      ]
    },
    "merge vertical 2": {
      "token": "merge_vertical_2",
      "merges": [
        {
          "token": "m_name",
          "column": "name"
        },
        {
          "token": "m_age",
          "column": "age"
        }
      ],
      "data": [
        {
          "name": "diep",
          "age": 23
        },
        {
          "name": "diep",
          "age": 34
        },
        {
          "name": "tung",
          "age": 23
        },
        {
          "name": "tung dx",
          "age": 23
        }
      ]
    }
  }
}
'
```