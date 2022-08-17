package com.zzzzzzzs.readword;

import cn.hutool.core.io.FileUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;
import java.util.List;

public class ReadWord {
  private static final Logger log = LoggerFactory.getLogger(ReadWord.class);
  // 读取 word 内容
  static String read(File file, InputStream is) throws Exception {
    ZipSecureFile.setMinInflateRatio(0);
    String text = "";
    if (FileMagic.valueOf(is) == FileMagic.OLE2) {
      WordExtractor ex = new WordExtractor(is);
      text = ex.getText();
      ex.close();
    } else if (FileMagic.valueOf(is) == FileMagic.OOXML) {
      XWPFDocument doc = new XWPFDocument(is);
      XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
      text = extractor.getText();
      extractor.close();
    }
    return text;
  }

  public static void main(String[] args) {
    if (args.length < 1) {
      throw new IllegalArgumentException("请输入路径");
    }
    List<File> fileList = FileUtil.loopFiles(args[0]);

    fileList.parallelStream()
        .forEach(
            file -> {
              // 获取文件名后缀
              InputStream is = FileUtil.getInputStream(file);
              if (!FileUtil.getSuffix(file).equals("docx")
                  && !FileUtil.getSuffix(file).equals("doc")) {
                return;
              }
              try {
                String read = read(file, is);

                if (StringUtils.isNotBlank(read)) {
                  // 在 read 每行结尾加 file
                  String[] split = read.split("\n");
                  StringBuilder sb = new StringBuilder();
                  for (String s : split) {
                    sb.append(s).append("file name:" + file).append("\n");
                  }
                  log.info(sb.toString());
                }
              } catch (Exception e) {
                e.printStackTrace();
              }
            });
  }
}
