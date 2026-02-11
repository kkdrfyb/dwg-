class DxfEntityText {
  const DxfEntityText({
    required this.type,
    required this.layer,
    required this.text,
  });

  final String type;
  final String layer;
  final String text;
}

class DxfSearchResult {
  const DxfSearchResult({
    required this.fileName,
    required this.objectType,
    required this.layer,
    required this.keyword,
    required this.content,
  });

  final String fileName;
  final String objectType;
  final String layer;
  final String keyword;
  final String content;
}

class DxfValveResult {
  const DxfValveResult({
    required this.fileName,
    required this.kind,
    required this.name,
    required this.code,
    required this.size,
    required this.height,
  });

  final String fileName;
  final String kind;
  final String name;
  final String code;
  final String size;
  final String height;
}

class DxfReplacePair {
  DxfReplacePair({required this.find, required this.replaceWith});

  String find;
  String replaceWith;

  bool get isValid => find.trim().isNotEmpty;
}

class DxfReplaceResult {
  DxfReplaceResult({
    required this.fileName,
    required this.filePath,
    required this.objectType,
    required this.layer,
    required this.originalText,
    required this.updatedText,
    required this.rule,
    this.skip = false,
  });

  final String fileName;
  final String filePath;
  final String objectType;
  final String layer;
  final String originalText;
  final String updatedText;
  final String rule;
  bool skip;
}
