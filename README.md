<div id="table-of-contents">
<h2>Table of Contents</h2>
<div id="text-table-of-contents">
<ul>
<li><a href="#sec-1">1. xlsx-json</a>
<ul>
<li><a href="#sec-1-1">1.1. Installation</a></li>
<li><a href="#sec-1-2">1.2. Usage</a></li>
<li><a href="#sec-1-3">1.3. Config</a></li>
<li><a href="#sec-1-4">1.4. Rules</a></li>
<li><a href="#sec-1-5">1.5. Use Cases Is King</a></li>
</ul>
</li>
</ul>
</div>
</div>

# xlsx-json<a id="sec-1" name="sec-1"></a>

## Installation<a id="sec-1-1" name="sec-1-1"></a>

    $ npm install xlsx-json

## Usage<a id="sec-1-2" name="sec-1-2"></a>

    var xlsx2json = require('xlsx-json');
    
    var config = require('./task.json');
    xlsx2json(config, function(err) {
        if (err) {
            console.log(err);
        }
    });

## Config<a id="sec-1-3" name="sec-1-3"></a>

config work in `task.json`.
`task.json` is an array of work object,
each work object contains `input` , `sheet` , `range`, `output` the four properties. It's like

    [
        {
            "input": "xxx.xlsx",
            "sheet": "sheetName1",
            "range": "A1:H52",
            "output": "xxxx.json"
        }
    ]

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
    
    <table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">
    
    
    <colgroup>
    <col  class="left" />
    
    <col  class="left" />
    
    <col  class="right" />
    </colgroup>
    <tbody>
    <tr>
    <td class="left">&#xa0;</td>
    <td class="left">enabel</td>
    <td class="right">actions</td>
    </tr>
    
    
    <tr>
    <td class="left">battle.pve</td>
    <td class="left">true</td>
    <td class="right">1</td>
    </tr>
    
    
    <tr>
    <td class="left">battle.pvp</td>
    <td class="left">false</td>
    <td class="right">2</td>
    </tr>
    </tbody>
    </table>
    
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

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="right" />

<col  class="left" />

<col  class="left" />

<col  class="right" />

<col  class="left" />
</colgroup>
<tbody>
<tr>
<td class="left">atk</td>
<td class="right">100</td>
<td class="left">def</td>
<td class="left">200</td>
<td class="right">&#xa0;</td>
<td class="left">&#xa0;</td>
</tr>


<tr>
<td class="left">speed</td>
<td class="right">300</td>
<td class="left">&#xa0;</td>
<td class="left">power</td>
<td class="right">400</td>
<td class="left">&#xa0;</td>
</tr>


<tr>
<td class="left">hp</td>
<td class="right">500</td>
<td class="left">magic</td>
<td class="left">&#xa0;</td>
<td class="right">600</td>
<td class="left">700</td>
</tr>
</tbody>
</table>

to

    {
        "600": 700,
        "atk": 100,
        "def": 200,
        "speed": 300,
        "power": 400,
        "hp": 500
    }

When `test` is undefined

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="right" />
</colgroup>
<tbody>
<tr>
<td class="left">test.a.b</td>
<td class="right">120</td>
</tr>
</tbody>
</table>

to

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

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="left" />

<col  class="right" />

<col  class="right" />

<col  class="right" />

<col  class="right" />
</colgroup>
<tbody>
<tr>
<td class="left">testArr</td>
<td class="left">[]</td>
<td class="right">1</td>
<td class="right">2</td>
<td class="right">3</td>
<td class="right">4</td>
</tr>
</tbody>
</table>

to

    {
        "testArr": [
            1,
            2,
            3,
            4
        ]
    }

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="left" />

<col  class="right" />

<col  class="left" />
</colgroup>
<tbody>
<tr>
<td class="left">&#xa0;</td>
<td class="left">enabel</td>
<td class="right">actions</td>
<td class="left">vipRequired</td>
</tr>


<tr>
<td class="left">battle.pve</td>
<td class="left">true</td>
<td class="right">1</td>
<td class="left">false</td>
</tr>


<tr>
<td class="left">battle.pvp</td>
<td class="left">false</td>
<td class="right">2</td>
<td class="left">false</td>
</tr>


<tr>
<td class="left">battle.boss</td>
<td class="left">false</td>
<td class="right">3</td>
<td class="left">true</td>
</tr>


<tr>
<td class="left">battle.team</td>
<td class="left">true</td>
<td class="right">4</td>
<td class="left">true</td>
</tr>
</tbody>
</table>

to

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

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="right" />

<col  class="right" />

<col  class="right" />

<col  class="right" />

<col  class="right" />
</colgroup>
<tbody>
<tr>
<td class="left">&#xa0;</td>
<td class="right">card.S</td>
<td class="right">card.A</td>
<td class="right">card.B</td>
<td class="right">vip</td>
<td class="right">bonus</td>
</tr>


<tr>
<td class="left">rewards.1</td>
<td class="right">900</td>
<td class="right">600</td>
<td class="right">450</td>
<td class="right">3</td>
<td class="right">8</td>
</tr>


<tr>
<td class="left">rewards.2</td>
<td class="right">1200</td>
<td class="right">800</td>
<td class="right">600</td>
<td class="right">5</td>
<td class="right">16</td>
</tr>


<tr>
<td class="left">rewards.3</td>
<td class="right">1800</td>
<td class="right">1200</td>
<td class="right">900</td>
<td class="right">7</td>
<td class="right">24</td>
</tr>
</tbody>
</table>

to

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

<table border="2" cellspacing="0" cellpadding="6" rules="groups" frame="hsides">


<colgroup>
<col  class="left" />

<col  class="right" />

<col  class="left" />

<col  class="right" />

<col  class="right" />

<col  class="left" />

<col  class="right" />
</colgroup>
<tbody>
<tr>
<td class="left">&#xa0;</td>
<td class="right">.id</td>
<td class="left">.type</td>
<td class="right">.amount</td>
<td class="right">.id</td>
<td class="left">.type</td>
<td class="right">.amount</td>
</tr>


<tr>
<td class="left">rewards</td>
<td class="right">1001</td>
<td class="left">item</td>
<td class="right">50</td>
<td class="right">2001</td>
<td class="left">equip</td>
<td class="right">5</td>
</tr>


<tr>
<td class="left">rewards</td>
<td class="right">1002</td>
<td class="left">item</td>
<td class="right">100</td>
<td class="right">2002</td>
<td class="left">equip</td>
<td class="right">10</td>
</tr>
</tbody>
</table>

to

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

<div id="footnotes">
<h2 class="footnotes">Footnotes: </h2>
<div id="text-footnotes">

<div class="footdef"><sup><a id="fn.1" name="fn.1" class="footnum" href="#fnr.1">1</a></sup> <p>DEFINITION NOT FOUND.</p></div>

<div class="footdef"><sup><a id="fn.2" name="fn.2" class="footnum" href="#fnr.2">2</a></sup> <p>DEFINITION NOT FOUND.</p></div>

<div class="footdef"><sup><a id="fn.3" name="fn.3" class="footnum" href="#fnr.3">3</a></sup> <p>DEFINITION NOT FOUND.</p></div>


</div>
</div>
