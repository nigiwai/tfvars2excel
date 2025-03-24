# tfvars2excel

## 説明

tfvarsファイルとExcelファイルの相互変換を行います。また、Excelファイルからmap形式のtfvarsファイルを生成します。

## 前提条件

- Python 3.11以上
- pip (Pythonパッケージインストーラー)

## 準備

1. 仮想環境を作成してアクティブにします:

    ```sh
    python -m venv venv
    .\venv\Scripts\activate
    ```

2. 必要なパッケージをインストールします:

    ```sh
    pip install -r requirements.txt
    ```

3. `.env`ファイルを作成し、対象のシート名を指定します:

    ```env
    1_SHEET_NAME_PREFIXES='ヒアリングシート'
    2_BAN_WORDS=azurerm,General
    3_SHEET_NAME_PREFIXES=apcol,netcol,natcol,windows_vms,tags,ip_groups
    ```

## 使用方法

### Excelファイルからtfvarsファイルを生成

```cmd
python 1_excel2tfvars.py <excel_filepath>
```

### tfvarsファイルからExcelファイルを生成

```cmd
python 2_tfvars2excel.py <tfvars_filepath> <excel_filepath>
```

### Excel(2次元表)からmap出力

```cmd
python 3_excel2map.py <excel_filepath>
```

## ライセンス

MITライセンス
