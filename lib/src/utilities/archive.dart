part of '../../excel.dart';

Archive _cloneArchive(Archive archive, Map<String, ArchiveFile> archiveFiles, {String? excludedFile}) {
  final clone = Archive();
  for (final file in archive.files) {
    if (file.isFile) {
      if (excludedFile != null && file.name.toLowerCase() == excludedFile.toLowerCase()) {
        continue;
      }
      ArchiveFile copy;
      if (archiveFiles.containsKey(file.name)) {
        copy = archiveFiles[file.name]!;
      } else {
        final content = file.content;
        final compression = _noCompression.contains(file.name) ? CompressionType.none : CompressionType.deflate;
        copy = ArchiveFile(file.name, content.length, content)..compression = compression;
      }
      clone.addFile(copy);
    }
  }
  return clone;
}
