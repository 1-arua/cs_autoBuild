using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.IO.Compression;
using System.Data;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;

namespace AutoCompile
{
    class Batch
    {
        private string batchPath;

        public Batch(string bPath)  // bPath = batch path
        {
            batchPath = bPath;
        }

        public int callBatch()
        {
            Process process = Process.Start(batchPath);
            process.WaitForExit();
            return process.ExitCode;
        }
    }


    class Program
    {
        public static (string[] stuIDs, string[] stuName, string[] fileName) getStuFileName(string csvPath)
        {
            var parser = new TextFieldParser(csvPath)
            {
                TextFieldType = FieldType.Delimited,
                Delimiters = new string[] { "," }
            };

            List<string[]> rows = new List<string[]>();
            bool flag = false;
            while (!parser.EndOfData)
            {
                string[] temp = parser.ReadFields();
                if (temp[0] == "id")
                {
                    flag = true;
                }

                if (flag)
                {
                    rows.Add(temp);
                }
            }

            // 列設定
            DataTable table = new DataTable();
            table.Columns.AddRange(rows.First().Select(s => new DataColumn(s)).ToArray());

            // 行追加
            foreach (var row in rows.Skip(1))
            {
                table.Rows.Add(row);
            }

            List<string> fileNames_tmp = new List<string>();
            List<string> stuNames_tmp = new List<string>();
            List<string> stuIDs_tmp = new List<string>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                fileNames_tmp.Add(table.Rows[i]["download_report_file"].ToString());
                stuNames_tmp.Add(table.Rows[i]["氏名"].ToString());
                stuIDs_tmp.Add(table.Rows[i]["USERID"].ToString());
            }

            string[] fileNames = fileNames_tmp.ToArray();
            string[] stuNames = stuNames_tmp.ToArray();
            string[] stuIDs = stuIDs_tmp.ToArray();

            (string[] stuIDs, string[] stuName, string[] fileName) result = (stuIDs, stuNames, fileNames);
            return result;
        }

        static void makeExeShortcut(string targetPath, string shortcutPath)
        {
            IWshRuntimeLibrary.WshShell objShell = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.WshShortcut objShortcut = objShell.CreateShortcut(shortcutPath);
            objShortcut.TargetPath = targetPath;
            objShortcut.Save();
        }


        static void Main(string[] args)
        {
            Directory.SetCurrentDirectory(Path.GetDirectoryName(Environment.GetCommandLineArgs()[0]));
            if (Path.GetFileName(Directory.GetCurrentDirectory())=="Debug")
            {
                Directory.SetCurrentDirectory("../../..");  // バッチファイルのあるディレクトリに移動
            }
            else if (Path.GetFileName(Directory.GetCurrentDirectory()).Contains("AutoCompile"))
            {

            }
            else
            {
                ApplicationException e = new ApplicationException("実行ファイルがなんらかの理由でおかしいです．");
                throw e;
            }

            string batchPath = Path.Combine(Directory.GetCurrentDirectory(), "autoCompile.bat"); // バッチファイルのパス取得
            Batch batch = new Batch(batchPath);

            string zipPath = Console.ReadLine();  // 演習問題ごとのフォルダのzipのパスの入力
            if ((zipPath.Substring(0, 1) == "\"") & (zipPath.Substring(zipPath.Length - 1, 1) == "\""))
            {
                zipPath = zipPath.Substring(1, zipPath.Length - 2);
            }
            Directory.SetCurrentDirectory(Path.GetDirectoryName(zipPath));  // zipの親フォルダをカレントに
            if ((File.GetAttributes(zipPath)&FileAttributes.Directory) != FileAttributes.Directory)
            {
                if (System.IO.Path.GetExtension(zipPath) != ".zip")
                {
                    throw new InvalidDataException("Input must be ZIP file.");
                }
                string extractPath = Path.GetFileNameWithoutExtension(zipPath); // 以下2行，zipの展開
                ZipFile.ExtractToDirectory(zipPath, extractPath);
            }

            string parentDir = Path.Combine(
                Path.GetDirectoryName(zipPath),
                Path.GetFileNameWithoutExtension(zipPath)
                );  // 解凍したフォルダのパス
            Directory.SetCurrentDirectory(parentDir);  // 解凍(入力)したディレクトリに移動

            Directory.CreateDirectory("_results");  // 学生の提出ファイルの展開先のディレクトリ生成

            var tmpTuple = getStuFileName("answer-utf8.txt");   // answer-utf8.txtをもとに受講者の学生ID，氏名，提出ファイル名を取得
            string[] stuNames = tmpTuple.stuName;
            string[] fileNames = tmpTuple.fileName;
            string[] stuIDs = tmpTuple.stuIDs;

            File.AppendAllText(Path.Combine("_results", "students with errors.log"), 
                "Student Name, Exception Type, Error Message" + Environment.NewLine, Encoding.UTF8);    // エラーログファイルの生成(ヘッダーのみ)

            for (int i= 0; i < fileNames.Length; i++)   // 学生ごとにビルド
            {
                string stuName = stuNames[i];
                string fileName = fileNames[i];
                string stuID = stuIDs[i];
                stuName = stuID + "_" + Regex.Replace(stuName, @"[\s]+", "");

                try     // エラーをすべてキャッチ，エラーログに
                {
                    if (Path.GetExtension(fileName) != ".zip")  // 提出ファイルがzipか確認
                    {
                        if (Path.GetExtension(fileName) == ".cs")   // csなら保留
                        {
                            throw new InvalidDataException("提出ファイルミス(.csなので採点可かもしれない)");
                        }
                        else                                    // slnとかなら採点不可(0点？)
                        {
                            throw new InvalidDataException("提出ファイルミス(zipではない，採点不可)");
                        }
                    }
                    string extractPath = "_results";    // 展開先ディレクトリ名
                    string folderName = stuName;        // 展開時に生成するフォルダの名前
                    ZipFile.ExtractToDirectory(fileName, Path.Combine(extractPath, folderName));    // 学生の提出zipファイルの展開

                    Directory.SetCurrentDirectory(Path.Combine(extractPath, folderName));   // 展開したフォルダに移動

                    /*=====================================================================================
                    提出されたzipを展開し，要求されたディレクトリ構造になっているかチェック(洗練の余地あり)
                    =====================================================================================*/
                    string[] subDirs = Directory.GetDirectories(Directory.GetCurrentDirectory());
                    // Console.WriteLine(String.Join(", ", subDirs));
                    string[] subFiles = Directory.GetFiles(Directory.GetCurrentDirectory());
                    string dirName = Path.GetFileNameWithoutExtension(subDirs[0]);
                    if ((subDirs.Length == 1) &
                            File.Exists(Path.Combine(dirName, dirName, dirName + ".csproj")))
                    {
                        Directory.SetCurrentDirectory(Path.Combine(dirName, dirName));
                    }
                    else if ((subDirs.Length == 1) &
                            File.Exists(Path.Combine(dirName, dirName + ".csproj")))
                    {
                        Directory.SetCurrentDirectory(dirName);
                    }
                    else if (File.Exists(Path.Combine(dirName + ".csproj")))
                    {

                    }
                    else
                    {
                        throw new InvalidDataException("提出ファイルミス");
                    }
                    /*=====================================================================================
                    ディレクトリ構造チェックここまで
                    =====================================================================================*/


                    // 学生がビルドして提出していた場合，元々存在する実行ファイルを.oldに変更
                    // ビルドし直したものがエラーの場合に古いものを一応確認できるように
                    if (File.Exists(Path.Combine("bin", "Debug", dirName + ".exe")))    
                    {
                        Microsoft.VisualBasic.FileSystem.Rename(
                            Path.Combine("bin", "Debug", dirName + ".exe"),
                            Path.Combine("bin", "Debug", dirName + ".exe.old"));
                    }

                    if (batch.callBatch() == 1)     // ビルドの実行．ビルドエラーの場合エラーコード1が返り，ifの中が実行
                    {
                        throw new InvalidDataException("コンパイルエラー");
                    }

                    // ショートカットの作成
                    string targetPath = "";
                    if (File.Exists(Path.Combine("bin", "Debug", dirName + ".exe")))
                    {
                        targetPath = Path.Combine(Directory.GetCurrentDirectory(), "bin", "Debug", dirName + ".exe");
                    }
                    else if (File.Exists(Path.Combine("bin", "Debug", "netcoreapp3.1", dirName + ".exe")))
                    {
                        targetPath = Path.Combine("bin", "Debug", "netcoreapp3.1", dirName + ".exe");
                        targetPath = Path.Combine(Directory.GetCurrentDirectory(), targetPath);
                    }
                    else
                    {
                        throw new FileNotFoundException("Compiled file is not found.");     // 多分発生しないエラー
                    }
                    string shortcutPath = Path.Combine(parentDir, extractPath, stuName + ".lnk");
                    makeExeShortcut(targetPath, shortcutPath);
                }
                catch (System.ComponentModel.Win32Exception e)      // batファイルが見つからなかった場合のエラー
                {
                    Directory.SetCurrentDirectory(parentDir);
                    string[] msg = { };
                    File.AppendAllText(Path.Combine("_results", "students with errors.log"),
                        stuName + ", " + e.GetType().Name + ":," + e.Message + Environment.NewLine,
                        Encoding.UTF8);

                    Console.WriteLine("Data");
                    Console.WriteLine(e.Data);
                    Console.WriteLine("InnerException");
                    Console.WriteLine(e.InnerException);
                    Console.WriteLine("Message");
                    Console.WriteLine(e.Message);
                    Console.WriteLine("NativeErrorCode");
                    Console.WriteLine(e.NativeErrorCode);
                    Console.WriteLine("Source");
                    Console.WriteLine(e.Source);
                    Console.WriteLine("StackTrace");
                    Console.WriteLine(e.StackTrace);
                    Console.WriteLine("TargetSite");
                    Console.WriteLine(e.TargetSite);
                }
                catch (Exception e)     // 上以外のすべてのエラー(エラーログに生徒名とともに残す)
                {
                    Directory.SetCurrentDirectory(parentDir);
                    string[] msg = { };
                    File.AppendAllText(Path.Combine("_results", "students with errors.log"),
                        stuName + ", " + e.GetType().Name + ":," + e.Message + Environment.NewLine,
                        Encoding.UTF8);
                }

                Directory.SetCurrentDirectory(parentDir);

            }

            Console.ReadLine();

        }
    }
}