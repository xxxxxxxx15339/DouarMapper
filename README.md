<div align="center">
<pre>
██████╗  ██████╗ ██╗   ██╗ █████╗ ██████╗     ███╗   ███╗ █████╗ ██████╗ ██████╗ ███████╗██████╗ 
██╔══██╗██╔═══██╗██║   ██║██╔══██╗██╔══██╗    ████╗ ████║██╔══██╗██╔══██╗██╔══██╗██╔════╝██╔══██╗
██║  ██║██║   ██║██║   ██║███████║██████╔╝    ██╔████╔██║███████║██████╔╝██████╔╝█████╗  ██████╔╝
██║  ██║██║   ██║██║   ██║██╔══██║██╔══██╗    ██║╚██╔╝██║██╔══██║██╔═══╝ ██╔═══╝ ██╔══╝  ██╔══██╗
██████╔╝╚██████╔╝╚██████╔╝██║  ██║██║  ██║    ██║ ╚═╝ ██║██║  ██║██║     ██║     ███████╗██║  ██║
╚═════╝  ╚═════╝  ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝    ╚═╝     ╚═╝╚═╝  ╚═╝╚═╝     ╚═╝     ╚══════╝╚═╝  ╚═╝
                                                                                                 
---------------------------------------------------
Python CLI tool to clean, normalize & deduplicate Moroccan village lists
</pre>

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Python Version](https://img.shields.io/badge/python-%3E%3D3.7-blue.svg)](https://www.python.org/)
</div>

Is Excel cleaning too tedious? Want something made for humans? Try **DouarMapper**—a Python CLI that reads your `.xlsx` file, normalizes text, merges similar names, and spits out a neat, formatted workbook.

## Installation

Clone this repo and install dependencies.  
*(Requires Python 3.7+)*

```sh
git clone https://github.com/YasserBAOUZIL/DouarMapper.git
cd DouarMapper
python3 -m venv venv
source venv/bin/activate      # On Windows: venv\Scripts\activate
pip install pandas openpyxl
```

## Usage

1. **Place your Excel file** (e.g. `Douars.xlsx`) in the project root.
2. **Run the CLI**:

   ```sh
   python DouarMapper.py
   ```
3. **Follow the prompts** to enter:

   * **Starting row**
   * **Ending row**
   * **Starting column** (A = 1, B = 2, …)
   * **Ending column**

The script will then:

1. Print an initial **committee → douar** mapping.
2. Prompt to merge similar committee names (default ≥ 80% similarity).
3. Prompt to merge similar douar names within each committee.
4. Export `committees_output.xlsx` with two sheets:

   * **Cleaned Communes** (one row per Commune – Douar)
   * **Probable Duplicates** (groups of similar douars)

### Example Session

```txt
$ python DouarMapper.py
Enter the starting row of the table:
> 10
Enter the ending row of the table:
> 250
Enter the starting column of the table  (A=1, B=2 ...):
> 1
Enter the ending column of the table  (A=1, B=2 ...):
> 3

Initial committee→douar mapping:
Committee: Marrakech
  - Douar A
  - Douar B

These committees are similar (>80%):
1. Agadir
2. Agadir-Idao  
Are these the same committee? (y/n): y  
Which name do you want to keep? Enter the number:
1. Agadir
2. Agadir-Idao  
Your choice: 1

# …then douar merging prompts…

Excel file 'committees_output.xlsx' created and formatted.
```

## Configuration

* **Similarity threshold**
  Tweak the `threshold` parameter in `group_similar_items(..., threshold=0.8)` for stricter or looser matching.

* **Input/output filenames**
  Edit the default `'Douars.xlsx'` and `'committees_output.xlsx'` in the script if needed.

## IO Redirection (Optional)

Capture console output to a file:

```sh
python DouarMapper.py > run.log 2>&1
```

## Development

1. Fork this repo
2. Create a feature branch

   ```sh
   git checkout -b feature/my-change
   ```
3. Install requirements

   ```sh
   pip install -r requirements.txt
   ```
4. Make your changes & commit

   ```sh
   git commit -am "Add awesome feature"
   ```
5. Push & open a Pull Request

   ```sh
   git push origin feature/my-change
   ```

## License

Distributed under the MIT License. See [LICENSE](./LICENSE) for details.

## Author

**Yasser BAOUZIL** – [GitHub](https://github.com/xxxxxxxx15339)

