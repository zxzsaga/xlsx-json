# xlsx-json<a id="sec-1" name="sec-1"></a>

## Installation<a id="sec-1-1" name="sec-1-1"></a>

    $ npm install xlsx-json

## Usage<a id="sec-1-2" name="sec-1-2"></a>

    var xlsx2json = require('xlsx-json');
    
    var task = require('./task.json');
    xlsx2json(task, function(err, jsonArr) {
        if (err) {
            console.log(err);
            return;
        }
        // Do sth with jsonArr
    });

## Config<a id="sec-1-3" name="sec-1-3"></a>

config work in `task.json`.
`task.json` is an array of work object,
each work object contains `input` , `sheet` , `range`, `raw`, `output` the four properties. It's like

    [
        {
            "input": "xxx.xlsx",
            "sheet": "sheetName1",
            "range": "A1:H52",
            "raw": true,
            "output": "xxxx.json"
        },
        {
            "input": "yyy.xlsx",
            "sheet": "sheetName2"
        }
    ]

-   `input` and `sheet` are necessary, `range`, `raw` and `output` are optional.
-   If `range` is not set, range will be the `!ref` range of the sheet. You can use this option to limit the range when you may write some comment in somewhere and you don't want to parse them.
-   If `raw` is set to `true`, the sheet will be parsed to a two dimension array like the raw xlsx table.
-   If `output` is not set, it won't write file. This option is just for convenience and you can handle the `jsonArr` and write it to files by your own.

## Rules<a id="sec-1-4" name="sec-1-4"></a>

-   Root is an object, not an array.
-   Parse assign range of file. Read file from top to bottom, from left to right.
-   Keys will be read like the properties in JSON.
    For example, `test.a[1][2].b` represent what it is in JSON.
    If the properties don't exist, they will be generated.
    So if the root doesn't have the property `test` before, it will be
    
        {
            "test": {
                "a": [
                    null,
                    [
                        null,
                        null,
                        {
                            "b": 120
                        }
                    ]
                ]
            }
        }

-   In normal cases, keys and values come in pairs.
    A non-empty value represent a key, and the element in next column is the value.
    You can write only one key-value pair in a row, you can also write several relative key-value pairs in a row if you like.
-   If the first element of a row is empty and the second is not empty, it represent the beginning of a table.
    The elements in this row after the empty cell are the table head.
    From the next row, the concat of first element and the table head represent the key, the element in corresponding cell is the value.
    
        |            | enabel | actions |
        | battle.pve | true   |       1 |
        | battle.pvp | false  |       2 |

    will be parsed to
    
        {
            "battle": {
                "pve": {
                    "enable": true,
                    "actions": 1
                },
                "pvp": {
                    "enable": false,
                    "actions": 2
                }
            }
        }

-   If first two elements of a row are both empty, they will end a table.
    So you can use an empty row to end a table.
-   If a value if `[]` , the next several consecutive elements will be pushed to this empty array.

## Use Cases Is King<a id="sec-1-5" name="sec-1-5"></a>

    | atk   | 100 | def   |   200 |     |     |
    | speed | 300 |       | power | 400 |     |
    | hp    | 500 | magic |       | 600 | 700 |

will be parsed to

    {
        "600": 700,
        "atk": 100,
        "def": 200,
        "speed": 300,
        "power": 400,
        "hp": 500
    }


If the `raw` option is set to `true`, the above table will be parsed to

    [
        [
            "atk",
            100,
            "def",
            200,
            null,
            null
        ],
        [
            "speed",
            300,
            null,
            "power",
            400,
            null
        ],
        [
            "hp",
            500,
            "magic",
            null,
            600,
            700
        ]
    ]

When `test` is undefined,

    | test.a[1][2].b | 120 |

will be parsed to

    {
        "test": {
            "a": [
                null,
                [
                    null,
                    null,
                    {
                        "b": 120
                    }
                ]
            ]
        }
    }

Array:

    | testArr | [] | 1 | 2 | 3 | 4 |

will be parsed to

    {
        "testArr": [
            1,
            2,
            3,
            4
        ]
    }

Key concat:

    |                 | enabel | actions | vipRequired |
    | battle.pve      | true   |       1 | false       |
    | battle.pvp      | false  |       2 | false       |
    | battle.boss     | false  |       3 | true        |
    | battle.team     | true   |       4 | true        |

will be parsed to

    {
        "battle": {
            "pve": {
                "enable": true,
                "actions": 1,
                "vipRequired" false
            },
            "pvp": {
                "enable": false,
                "actions": 2,
                "vipRequired" false
            },
            "boss": {
                "enable": false,
                "actions": 3,
                "vipRequired" true
            },
            "team": {
                "enable": true,
                "actions": 4,
                "vipRequired" true
            }
        }
    }

More complex:

    |           | card.S | card.A | card.B | vip | bonus |
    | rewards.1 |    900 |    600 |    450 |   3 |     8 |
    | rewards.2 |   1200 |    800 |    600 |   5 |    16 |
    | rewards.3 |   1800 |   1200 |    900 |   7 |    24 |

will be parsed to

    {
        "rewards": {
            "1": {
                "card": {
                    "S": 900,
                    "A": 600,
                    "B": 450
                },
                "vip": 3,
                "bonus": 8
            },
            "2": {
                "card": {
                    "S": 1200,
                    "A": 800,
                    "B": 600
                },
                "vip": 5,
                "bonus": 16
            },
            "3": {
                "card": {
                    "S": 1800,
                    "A": 1200,
                    "B": 900
                },
                "vip": 7,
                "bonus": 24
            }
        }
    }

Key contains array:

    |            | [0].id | [0].type | [0].amount | [1].id | [1].type | [1].amount |
    | rewards[0] |   1001 | item     |         50 |   2001 | equip    |          5 |
    | rewards[1] |   1002 | item     |        100 |   2002 | equip    |         10 |

will parsed to

    {
        "rewards": [
            [
                {
                    "id": 1001,
                    "type": "item",
                    "amount": 50
                },
                {
                    "id": 2001,
                    "type": "equip",
                    "amount": 5
                }
            ],
            [
                {
                    "id": 1002,
                    "type": "item",
                    "amount": 100
                },
                {
                    "id": 2002,
                    "type": "equip",
                    "amount": 10
                }
            ]
        ]
    }

**It's convenient and flexible, isn't it ?**
