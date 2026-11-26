// A2R Money — الأرصدة والمصروفات والسلف والفواتير (تليفونين، شهر بشهر)
// ملف واحد. البيانات كلها على الموبايل + نسخ احتياطي تلقائي في فولدر «A2R Money».
import 'dart:async';
import 'dart:convert';
import 'dart:io';
import 'dart:math';
import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:excel/excel.dart' as xls;
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import 'package:share_plus/share_plus.dart';
import 'package:open_filex/open_filex.dart';
import 'package:archive/archive.dart';

const String kAppName = 'A2R Money';
const String kAppVersion =
    String.fromEnvironment('APP_VERSION', defaultValue: '1.0.0');
const String kPrefsKey = 'a2r_money_data_v1';
const String kFolder = 'A2R Money';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const A2RMoneyApp());
}

// ==================== الشكل (هوية A2R Money) ====================
// زيتي غامق + زمردي + دهبي — مختلف عن A2R Team

class C {
  static const bg = Color(0xFF041F1E);
  static const bg2 = Color(0xFF07302E);
  static const card = Color(0xFF0B3B38);
  static const card2 = Color(0xFF0F4A46);
  static const line = Color(0xFF1C5C57);
  static const emerald = Color(0xFF2DD4A3);
  static const gold = Color(0xFFF2C14E);
  static const goldDeep = Color(0xFFC9962B);
  static const text = Color(0xFFF3F7F5);
  static const muted = Color(0xFF9CC3BC);
  static const red = Color(0xFFFF6B6B);
  static const blue = Color(0xFF67C7F5);
  static const violet = Color(0xFFB79CFF);
}

class A2RMoneyApp extends StatelessWidget {
  const A2RMoneyApp({super.key});

  @override
  Widget build(BuildContext context) {
    final base = ThemeData(useMaterial3: true, brightness: Brightness.dark);
    return MaterialApp(
      title: kAppName,
      debugShowCheckedModeBanner: false,
      theme: base.copyWith(
        scaffoldBackgroundColor: C.bg,
        colorScheme: const ColorScheme.dark(
          primary: C.gold,
          onPrimary: Color(0xFF231A00),
          secondary: C.emerald,
          onSecondary: Color(0xFF00241A),
          surface: C.card,
          onSurface: C.text,
          error: C.red,
        ),
        appBarTheme: AppBarTheme(
          backgroundColor: C.bg,
          foregroundColor: C.text,
          elevation: 0,
          centerTitle: false,
          titleTextStyle: base.textTheme.titleLarge?.copyWith(
              fontSize: 20, fontWeight: FontWeight.w800, color: C.text),
        ),
        cardTheme: CardThemeData(
          color: C.card,
          elevation: 0,
          margin: EdgeInsets.zero,
          shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(20),
              side: const BorderSide(color: C.line)),
        ),
        inputDecorationTheme: InputDecorationTheme(
          filled: true,
          fillColor: C.bg2,
          labelStyle: const TextStyle(color: C.muted),
          border: OutlineInputBorder(
              borderRadius: BorderRadius.circular(14),
              borderSide: const BorderSide(color: C.line)),
          enabledBorder: OutlineInputBorder(
              borderRadius: BorderRadius.circular(14),
              borderSide: const BorderSide(color: C.line)),
          focusedBorder: OutlineInputBorder(
              borderRadius: BorderRadius.circular(14),
              borderSide: const BorderSide(color: C.gold, width: 1.6)),
        ),
        navigationBarTheme: NavigationBarThemeData(
          backgroundColor: C.bg2,
          indicatorColor: C.gold.withValues(alpha: 0.22),
          labelTextStyle: WidgetStateProperty.all(
              const TextStyle(fontSize: 11.5, fontWeight: FontWeight.w600)),
        ),
        snackBarTheme: const SnackBarThemeData(
            backgroundColor: C.card2,
            contentTextStyle: TextStyle(color: C.text)),
        dialogTheme: DialogThemeData(
          backgroundColor: C.bg2,
          shape:
              RoundedRectangleBorder(borderRadius: BorderRadius.circular(22)),
        ),
        bottomSheetTheme: const BottomSheetThemeData(
          backgroundColor: C.bg2,
          shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.vertical(top: Radius.circular(26))),
        ),
      ),
      builder: (context, child) =>
          Directionality(textDirection: TextDirection.rtl, child: child!),
      home: const RootScreen(),
    );
  }
}

// ==================== أدوات ====================

const List<String> arabicMonths = [
  'يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو',
  'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'
];
const List<String> arabicWeekdays = [
  'الاتنين', 'التلات', 'الأربع', 'الخميس', 'الجمعة', 'السبت', 'الحد'
];

String two(int n) => n.toString().padLeft(2, '0');
String isoDate(DateTime d) => '${d.year}-${two(d.month)}-${two(d.day)}';
String monthKeyOf(DateTime d) => '${d.year}-${two(d.month)}';
DateTime parseIso(String s) => DateTime.tryParse(s) ?? DateTime.now();
String arDate(DateTime d) => '${d.day} ${arabicMonths[d.month - 1]} ${d.year}';
String arDateShort(String iso) {
  final d = parseIso(iso);
  return '${arabicWeekdays[d.weekday - 1]} ${d.day}/${d.month}';
}

/// عدد أيام الشهر (فبراير ٢٨/٢٩، و٣٠ أو ٣١).
int daysInMonth(int year, int month) => DateTime(year, month + 1, 0).day;
DateTime monthFirst(String mk) =>
    DateTime(int.parse(mk.substring(0, 4)), int.parse(mk.substring(5, 7)), 1);
DateTime monthLast(String mk) {
  final f = monthFirst(mk);
  return DateTime(f.year, f.month, daysInMonth(f.year, f.month));
}

String shiftMonth(String mk, int delta) {
  final f = monthFirst(mk);
  return monthKeyOf(DateTime(f.year, f.month + delta, 1));
}

String monthLabel(String mk) {
  final f = monthFirst(mk);
  return '${arabicMonths[f.month - 1]} ${f.year}';
}

bool inMonth(String iso, String mk) =>
    iso.length >= 7 && iso.substring(0, 7) == mk;

/// الفلوس متخزنة بالقروش (int) عشان الحساب ميبقاش فيه كسور غلط.
String money(int p, {bool sign = false}) {
  final neg = p < 0;
  final a = p.abs();
  final pounds = a ~/ 100;
  final pi = a % 100;
  final s = pounds.toString().replaceAllMapped(
      RegExp(r'(\d)(?=(\d{3})+$)'), (m) => '${m[1]},');
  final body = pi == 0 ? s : '$s.${two(pi)}';
  return '${neg ? '-' : (sign && p > 0 ? '+' : '')}$body';
}

String egp(int p, {bool sign = false}) => '${money(p, sign: sign)} جنيه';

/// يقبل أرقام عربي (٠١٢) وفارسي والفاصلة «٫» أو «,».
int? parseMoney(String raw) {
  var s = raw.trim();
  const ar = '٠١٢٣٤٥٦٧٨٩';
  const fa = '۰۱۲۳۴۵۶۷۸۹';
  final b = StringBuffer();
  for (final ch in s.split('')) {
    final i = ar.indexOf(ch);
    final j = fa.indexOf(ch);
    if (i >= 0) {
      b.write(i);
    } else if (j >= 0) {
      b.write(j);
    } else if (ch == '٫') {
      b.write('.');
    } else if (ch == ',' || ch == '،') {
      b.write(',');
    } else if (ch == ' ' || ch == '٬') {
      // فاصل آلاف
    } else {
      b.write(ch);
    }
  }
  s = b.toString();
  // «1,500» = ألف وخمسمية (فاصل آلاف)، لكن «12,5» = اتناشر ونص
  if (RegExp(r'^\d{1,3}(,\d{3})+(\.\d+)?$').hasMatch(s)) {
    s = s.replaceAll(',', '');
  } else {
    s = s.replaceAll(',', '.');
  }
  if (s.isEmpty) return null;
  final v = double.tryParse(s);
  if (v == null || v.isNaN || v.isInfinite || v < 0) return null;
  return (v * 100).round();
}

String normalizeAr(String s) => s
    .trim()
    .toLowerCase()
    .replaceAll(RegExp('[أإآ]'), 'ا')
    .replaceAll('ى', 'ي')
    .replaceAll('ة', 'ه')
    .replaceAll(RegExp(r'\s+'), ' ');

/// تنضيف الكلام اللي بيتكتب بإيد: يشيل المسافات الزيادة (والمسافات الخفية)
String cleanLabel(String s) => s
    .replaceAll(RegExp('[\u200e\u200f\u200b\u0640\u00a0]'), ' ')
    .trim()
    .replaceAll(RegExp(r'\s+'), ' ');

/// يضيف كلمة للقايمة لو مش موجودة (من غير تكرار: بيتجاهل الهمزات/التاء المربوطة/الياء/المسافات/الكابيتال).
/// بيرجّع الشكل اللي اتحفظ (لو موجود قبل كده بيرجّع القديم زي ما هو).
String addChoice(List<String> list, String raw) {
  final c = cleanLabel(raw);
  if (c.isEmpty) return '';
  final n = normalizeAr(c);
  for (final x in list) {
    if (normalizeAr(x) == n) return x;
  }
  list.add(c);
  return c;
}

bool sameLabel(String a, String b) => normalizeAr(cleanLabel(a)) == normalizeAr(cleanLabel(b));

/// اقتراحات أسماء الفواتير في أول مرة (مش بيانات شخصية)
const List<String> kDefaultBillNames = ['كهربا', 'ميه', 'غاز', 'نت البيت', 'باقة الموبايل', 'إيجار'];

final Random _rnd = Random.secure();
String newId([String prefix = '']) {
  final t = DateTime.now().microsecondsSinceEpoch.toRadixString(36);
  final r = List.generate(6, (_) => _rnd.nextInt(36).toRadixString(36)).join();
  return '$prefix$t$r';
}

// ==================== البيانات ====================

const String kGroupPersonal = 'شخصي';
const String kGroupWork = 'شغل';
const String kGroupHome = 'بيت';
const List<String> kGroups = [kGroupPersonal, kGroupWork, kGroupHome];

const Map<String, List<String>> kDefaultCategories = {
  kGroupPersonal: ['أكل وشرب', 'مواصلات', 'علاج ودوا', 'لبس', 'موبايل وكروت', 'خروج', 'تانية'],
  kGroupWork: ['مواصلات شغل', 'أكل في الشغل', 'أدوات وطباعة', 'تانية'],
  kGroupHome: ['طلبات البيت', 'خضار وفاكهة', 'لحمة وفراخ', 'مدارس ودروس', 'تصليحات', 'تانية'],
};

/// أنواع العمليات
const String tExpense = 'expense'; // مصروف (من مكان)
const String tAdd = 'add'; // إضافة رصيد (لمكان)
const String tWithdraw = 'withdraw'; // سحب (من مكان)
const String tTransfer = 'transfer'; // تحويل من مكان لمكان
const String tAdvance = 'advance'; // سلفة لشخص (من مكان)
const String tRepay = 'repay'; // رد سلفة (لمكان)
const String tBill = 'bill'; // دفع فاتورة (من مكان)
const String tToPhone = 'tophone'; // تحويل للتليفون التاني (من مكان) — مش صرف
const String tFromPhone = 'fromphone'; // استلام تحويل من التليفون الأساسي (لمكان) — مش دخل جديد

/// مصدر «إضافة رصيد» اللي بيتحسب استلام من الأساسي (لو اتسجل بإيد)
const String kFromMainSource = 'من التليفون التاني';
const String kFromMainSourceOld = 'من التليفون الأساسي'; // v2 (لسه بيتعرف عليه)
bool isPhoneSource(String note) => sameLabel(note, kFromMainSource) || sameLabel(note, kFromMainSourceOld);

const Map<String, String> kTypeLabels = {
  tExpense: 'مصروف',
  tAdd: 'إضافة رصيد',
  tWithdraw: 'سحب',
  tTransfer: 'تحويل',
  tAdvance: 'سلفة',
  tRepay: 'رد سلفة',
  tBill: 'فاتورة',
  tToPhone: 'تحويل لتليفون تاني',
  tFromPhone: 'استلام من تليفون تاني',
};

/// العمليات اللي الفلوس فيها «خارجة» من المكان
bool isOut(String type) =>
    type == tExpense ||
    type == tWithdraw ||
    type == tAdvance ||
    type == tBill ||
    type == tTransfer ||
    type == tToPhone;

const String kPlaceCash = 'cash';
const String kPlaceBank = 'bank';
const String kPlaceWallet = 'wallet';
const String kPlaceOther = 'other';

class Place {
  String id;
  String name;
  String kind; // cash | bank | wallet | other
  int opening; // رصيد البداية بالقروش
  bool archived;
  Place(
      {required this.id,
      required this.name,
      this.kind = kPlaceOther,
      this.opening = 0,
      this.archived = false});
  Map<String, dynamic> toJson() =>
      {'id': id, 'name': name, 'kind': kind, 'opening': opening, 'archived': archived};
  factory Place.fromJson(Map<String, dynamic> j) => Place(
        id: '${j['id']}',
        name: '${j['name'] ?? ''}',
        kind: '${j['kind'] ?? kPlaceOther}',
        opening: (j['opening'] as num?)?.toInt() ?? 0,
        archived: j['archived'] == true,
      );
}

class Person {
  String id;
  String name;
  bool family; // من أفراد الأسرة (بيظهر في تقرير الأسرة)
  bool child;
  Person({required this.id, required this.name, this.family = true, this.child = false});
  Map<String, dynamic> toJson() =>
      {'id': id, 'name': name, 'family': family, 'child': child};
  factory Person.fromJson(Map<String, dynamic> j) => Person(
        id: '${j['id']}',
        name: '${j['name'] ?? ''}',
        family: j['family'] != false,
        child: j['child'] == true,
      );
}

class Bill {
  String id;
  String name;
  String kind; // بيت | نت
  int dueDay; // يوم الاستحقاق في الشهر
  int usual; // المبلغ المعتاد (اختياري، 0 = مفيش)
  bool active;
  Bill(
      {required this.id,
      required this.name,
      this.kind = 'بيت',
      this.dueDay = 1,
      this.usual = 0,
      this.active = true});
  Map<String, dynamic> toJson() => {
        'id': id,
        'name': name,
        'kind': kind,
        'dueDay': dueDay,
        'usual': usual,
        'active': active
      };
  factory Bill.fromJson(Map<String, dynamic> j) => Bill(
        id: '${j['id']}',
        name: '${j['name'] ?? ''}',
        kind: '${j['kind'] ?? 'بيت'}',
        dueDay: ((j['dueDay'] as num?)?.toInt() ?? 1).clamp(1, 31),
        usual: (j['usual'] as num?)?.toInt() ?? 0,
        active: j['active'] != false,
      );

  /// يوم الاستحقاق الفعلي في شهر معين (٣١ في فبراير ← آخر يوم في الشهر)
  DateTime dueIn(String mk) {
    final f = monthFirst(mk);
    return DateTime(f.year, f.month, min(dueDay, daysInMonth(f.year, f.month)));
  }
}

class Txn {
  String id;
  String type;
  int amount; // بالقروش، دايمًا موجب
  String date; // yyyy-mm-dd
  String placeId; // مكان الفلوس: «منين» للخارج، «فين» للداخل
  String toPlaceId; // للتحويل بس
  String group; // شخصي/شغل/بيت (مصروف، فاتورة، وسحب لو اختار)
  String category;
  String personId; // فرد الأسرة (مصروف) أو الشخص (سلفة/رد)
  String billId;
  String billMonth;
  String note; // السبب/المصدر (اختياري)
  String refId; // استلام تحويل: رقم التحويل اللي جه من الأساسي (عشان مايتسجلش مرتين)
  String createdAt;
  Txn({
    required this.id,
    required this.type,
    required this.amount,
    required this.date,
    required this.placeId,
    this.toPlaceId = '',
    this.group = '',
    this.category = '',
    this.personId = '',
    this.billId = '',
    this.billMonth = '',
    this.note = '',
    this.refId = '',
    String? createdAt,
  }) : createdAt = createdAt ?? DateTime.now().toIso8601String();

  Map<String, dynamic> toJson() => {
        'id': id,
        'type': type,
        'amount': amount,
        'date': date,
        'placeId': placeId,
        if (toPlaceId.isNotEmpty) 'toPlaceId': toPlaceId,
        if (group.isNotEmpty) 'group': group,
        if (category.isNotEmpty) 'category': category,
        if (personId.isNotEmpty) 'personId': personId,
        if (billId.isNotEmpty) 'billId': billId,
        if (billMonth.isNotEmpty) 'billMonth': billMonth,
        if (note.isNotEmpty) 'note': note,
        if (refId.isNotEmpty) 'refId': refId,
        'createdAt': createdAt,
      };
  factory Txn.fromJson(Map<String, dynamic> j) => Txn(
        id: '${j['id']}',
        type: '${j['type']}',
        amount: (j['amount'] as num?)?.toInt().abs() ?? 0,
        date: '${j['date']}',
        placeId: '${j['placeId'] ?? ''}',
        toPlaceId: '${j['toPlaceId'] ?? ''}',
        group: '${j['group'] ?? ''}',
        category: '${j['category'] ?? ''}',
        personId: '${j['personId'] ?? ''}',
        billId: '${j['billId'] ?? ''}',
        billMonth: '${j['billMonth'] ?? ''}',
        note: '${j['note'] ?? ''}',
        refId: '${j['refId'] ?? ''}',
        createdAt: '${j['createdAt'] ?? ''}',
      );
  Txn copy() => Txn.fromJson(toJson());

  /// يدخل في «إجمالي الصرف»: مصروف + فاتورة + سحب
  bool get isSpend => type == tExpense || type == tBill || type == tWithdraw;
}

/// شهر جاي من التليفون التاني (محفوظ لوحده، مش بيتخلط بأرصدة التليفون ده)
class ImportedMonth {
  final String phoneId;
  final String phoneName;
  final String month;
  final String exportedAt;
  final List<Txn> txns;
  final Map<String, String> placeNames;
  final Map<String, String> personNames;
  final Map<String, String> billNames;
  final Map<String, int> endBalances; // اسم المكان ← رصيده آخر الشهر
  ImportedMonth({
    required this.phoneId,
    required this.phoneName,
    required this.month,
    required this.exportedAt,
    required this.txns,
    required this.placeNames,
    required this.personNames,
    required this.billNames,
    required this.endBalances,
  });
  String get key => '$phoneId|$month';
  Map<String, dynamic> toJson() => {
        'phoneId': phoneId,
        'phoneName': phoneName,
        'month': month,
        'exportedAt': exportedAt,
        'txns': txns.map((e) => e.toJson()).toList(),
        'placeNames': placeNames,
        'personNames': personNames,
        'billNames': billNames,
        'endBalances': endBalances,
      };
  static Map<String, String> _sm(dynamic m) => m is Map
      ? m.map((k, v) => MapEntry('$k', '$v'))
      : <String, String>{};
  factory ImportedMonth.fromJson(Map<String, dynamic> j) => ImportedMonth(
        phoneId: '${j['phoneId']}',
        phoneName: '${j['phoneName'] ?? ''}',
        month: '${j['month']}',
        exportedAt: '${j['exportedAt'] ?? ''}',
        txns: ((j['txns'] as List?) ?? [])
            .whereType<Map>()
            .map((e) => Txn.fromJson(Map<String, dynamic>.from(e)))
            .toList(),
        placeNames: _sm(j['placeNames']),
        personNames: _sm(j['personNames']),
        billNames: _sm(j['billNames']),
        endBalances: (j['endBalances'] is Map)
            ? (j['endBalances'] as Map)
                .map((k, v) => MapEntry('$k', (v as num?)?.toInt() ?? 0))
            : <String, int>{},
      );
}

class TransferResult {
  final bool ok;
  final String message;
  final Txn? txn;
  const TransferResult(this.ok, this.message, {this.txn});
}

class ImportResult {
  final bool ok;
  final bool replaced;
  final String message;
  final ImportedMonth? month;
  const ImportResult(this.ok, this.message, {this.replaced = false, this.month});
}

/// ملخص شهر لتليفون واحد
/// نتيجة ملف القوائم (معاينة أو دمج)
class ListsResult {
  final bool ok;
  final String error;
  final int people, categories, bills, places, missingPlaces, other;
  final String fromName;
  const ListsResult(this.ok, this.error,
      {this.people = 0,
      this.categories = 0,
      this.bills = 0,
      this.places = 0,
      this.missingPlaces = 0,
      this.other = 0,
      this.fromName = ''});
  int get total => people + categories + bills + places + other;

  /// «5 أفراد، 12 بند، 3 فواتير…»
  String describe({bool withPlaces = true}) {
    String n(int c, String one, String two, String many) =>
        c == 1 ? one : (c == 2 ? two : '$c $many');
    final parts = <String>[
      if (people > 0) n(people, 'فرد واحد', 'فردين', 'أفراد'),
      if (categories > 0) n(categories, 'بند واحد', 'بندين', 'بنود'),
      if (bills > 0) n(bills, 'فاتورة واحدة', 'فاتورتين', 'فواتير'),
      if (withPlaces && places > 0) n(places, 'مكان واحد', 'مكانين', 'أماكن'),
      if (other > 0) n(other, 'اختيار واحد', 'اختيارين', 'اختيارات (مصادر/أسباب/أسماء فواتير)'),
    ];
    return parts.isEmpty ? 'مفيش حاجة جديدة' : parts.join('، ');
  }
}

class CopyResult {
  final bool ok;
  final String message;
  final int months, listsAdded;
  final String fromName, exportedAt;
  const CopyResult(this.ok, this.message, {this.months = 0, this.listsAdded = 0, this.fromName = '', this.exportedAt = ''});
}

/// اللي «تحديث» هيعمله
class SyncPlan {
  final List<Map<String, dynamic>> copies, months, lists, transfers;
  const SyncPlan({required this.copies, required this.months, required this.lists, required this.transfers});
  bool get isEmpty => copies.isEmpty && months.isEmpty && lists.isEmpty && transfers.isEmpty;
}

class MonthSummary {
  int expenses = 0; // مصروفات (من غير الفواتير)
  int bills = 0;
  int withdraws = 0;
  int added = 0;
  int advancesGiven = 0;
  int advancesBack = 0;
  int transfers = 0;
  int toPhone = 0; // اتحول للتليفون التاني (مش صرف)
  int fromPhone = 0; // اتاستلم من الأساسي (مش صرف ولا دخل جديد)
  final Map<String, int> byGroup = {};
  final Map<String, int> byCategory = {}; // «المجموعة • البند»
  final Map<String, int> byPerson = {}; // اسم الفرد ← المصروف
  final Map<String, int> byPlaceOut = {};
  int get totalSpend => expenses + bills + withdraws;
  int count = 0;

  void addAll(MonthSummary o) {
    expenses += o.expenses;
    bills += o.bills;
    withdraws += o.withdraws;
    added += o.added;
    advancesGiven += o.advancesGiven;
    advancesBack += o.advancesBack;
    transfers += o.transfers;
    toPhone += o.toPhone;
    fromPhone += o.fromPhone;
    count += o.count;
    void mrg(Map<String, int> a, Map<String, int> b) =>
        b.forEach((k, v) => a[k] = (a[k] ?? 0) + v);
    mrg(byGroup, o.byGroup);
    mrg(byCategory, o.byCategory);
    mrg(byPerson, o.byPerson);
    mrg(byPlaceOut, o.byPlaceOut);
  }
}

MonthSummary summarize(
  Iterable<Txn> txns,
  String mk, {
  required String Function(String id) personName,
  required String Function(String id) placeName,
}) {
  final s = MonthSummary();
  for (final t in txns) {
    if (!inMonth(t.date, mk)) continue;
    s.count++;
    switch (t.type) {
      case tExpense:
        s.expenses += t.amount;
        break;
      case tBill:
        s.bills += t.amount;
        break;
      case tWithdraw:
        s.withdraws += t.amount;
        break;
      case tAdd:
        // «إضافة رصيد» مصدرها «من التليفون الأساسي» = استلام تحويل، مش فلوس جديدة
        if (isPhoneSource(t.note)) {
          s.fromPhone += t.amount;
        } else {
          s.added += t.amount;
        }
        break;
      case tToPhone:
        s.toPhone += t.amount;
        break;
      case tFromPhone:
        s.fromPhone += t.amount;
        break;
      case tAdvance:
        s.advancesGiven += t.amount;
        break;
      case tRepay:
        s.advancesBack += t.amount;
        break;
      case tTransfer:
        s.transfers += t.amount;
        break;
    }
    if (t.isSpend) {
      final g = t.group.isNotEmpty
          ? t.group
          : t.type == tWithdraw
              ? 'سحب بدون بند'
              : kGroupHome;
      s.byGroup[g] = (s.byGroup[g] ?? 0) + t.amount;
      final cat = t.type == tBill
          ? 'فواتير • ${t.category.isEmpty ? 'فاتورة' : t.category}'
          : t.category.isNotEmpty
              ? '$g • ${t.category}'
              : t.type == tWithdraw
                  ? 'سحب بدون بند'
                  : '$g • بدون بند';
      s.byCategory[cat] = (s.byCategory[cat] ?? 0) + t.amount;
      if (t.personId.isNotEmpty) {
        final n = personName(t.personId);
        s.byPerson[n] = (s.byPerson[n] ?? 0) + t.amount;
      }
      final pn = placeName(t.placeId);
      s.byPlaceOut[pn] = (s.byPlaceOut[pn] ?? 0) + t.amount;
    }
  }
  return s;
}

class MoneyData {
  String phoneId;
  String phoneName;
  String role; // main | second | '' (لسه متحددش)
  List<Place> places;
  List<Person> people;
  List<Bill> bills;
  List<Txn> txns;
  Map<String, List<String>> categories;
  Map<String, ImportedMonth> imported;
  Map<String, String> copySeen; // تليفون ← وقت آخر نسخة كاملة اتاستلمت منه
  String lastPlaceId;
  String lastInPlaceId;
  List<String> sources; // مصادر الفلوس اللي بتدخل (إضافة رصيد)
  List<String> reasons; // أسباب السحب
  List<String> billNames; // أسماء الفواتير اللي اتكتبت قبل كده
  String peerName; // اسم التليفون التاني (من آخر ملف جه منه)
  String syncFolder; // فولدر الاستلام اللي اتحدد مرة واحدة ('' = لسه)

  MoneyData({
    required this.phoneId,
    this.phoneName = '',
    this.role = '',
    List<Place>? places,
    List<Person>? people,
    List<Bill>? bills,
    List<Txn>? txns,
    Map<String, List<String>>? categories,
    Map<String, ImportedMonth>? imported,
    Map<String, String>? copySeen,
    this.lastPlaceId = '',
    this.lastInPlaceId = '',
    List<String>? sources,
    List<String>? reasons,
    List<String>? billNames,
    this.peerName = '',
    this.syncFolder = '',
  })  : sources = sources ?? [],
        reasons = reasons ?? [],
        billNames = billNames ?? List<String>.from(kDefaultBillNames),
        places = places ?? [],
        people = people ?? [],
        bills = bills ?? [],
        txns = txns ?? [],
        categories = categories ??
            kDefaultCategories.map((k, v) => MapEntry(k, List<String>.from(v))),
        imported = imported ?? {},
        copySeen = copySeen ?? {};

  /// تسطيبة جديدة: فاضية (مفيش أي بيانات شخصية) — بس الأماكن التلاتة الأساسية برصيد صفر
  factory MoneyData.fresh() => MoneyData(
        phoneId: newId('ph'),
        places: [
          Place(id: newId('pl'), name: 'كاش', kind: kPlaceCash),
          Place(id: newId('pl'), name: 'بنك', kind: kPlaceBank),
          Place(id: newId('pl'), name: 'محفظة', kind: kPlaceWallet),
        ],
      );

  bool get isMain => role != 'second';

  /// اسم التليفون ده (لو فاضي ← الكلام القديم حسب الدور، للنسخ القديمة بس)
  String get badgeName {
    final n = cleanLabel(phoneName);
    return n.isEmpty ? (isMain ? 'الأساسي' : 'التليفون التاني') : n;
  }

  /// اسم التليفون التاني (اتعرف من آخر ملف جه منه)
  String get otherName {
    final n = cleanLabel(peerName);
    return n.isEmpty ? 'التليفون التاني' : n;
  }

  void learnPeer(dynamic name) {
    final n = cleanLabel('${name ?? ''}');
    if (n.isNotEmpty && n.length <= 60) peerName = n;
  }
  bool get hasAnyData =>
      txns.isNotEmpty || people.isNotEmpty || bills.isNotEmpty || imported.isNotEmpty ||
      places.any((p) => p.opening != 0);

  Map<String, dynamic> toJson() => {
        'app': kAppName,
        'format': 1,
        'phoneId': phoneId,
        'phoneName': phoneName,
        'role': role,
        'places': places.map((e) => e.toJson()).toList(),
        'people': people.map((e) => e.toJson()).toList(),
        'bills': bills.map((e) => e.toJson()).toList(),
        'txns': txns.map((e) => e.toJson()).toList(),
        'categories': categories,
        'imported': imported.map((k, v) => MapEntry(k, v.toJson())),
        'copySeen': copySeen,
        'lastPlaceId': lastPlaceId,
        'lastInPlaceId': lastInPlaceId,
        'sources': sources,
        'reasons': reasons,
        'billNames': billNames,
        'peerName': peerName,
        'syncFolder': syncFolder,
      };

  factory MoneyData.fromJson(Map<String, dynamic> j) {
    List<T> l<T>(String k, T Function(Map<String, dynamic>) f) =>
        ((j[k] as List?) ?? [])
            .whereType<Map>()
            .map((e) => f(Map<String, dynamic>.from(e)))
            .toList();
    List<String> sl(String k) {
      final out = <String>[];
      if (j[k] is List) {
        for (final e in j[k] as List) {
          addChoice(out, '$e');
        }
      }
      return out;
    }

    final cats = <String, List<String>>{};
    if (j['categories'] is Map) {
      (j['categories'] as Map).forEach((k, v) {
        if (v is List) {
          final l = <String>[];
          for (final e in v) {
            addChoice(l, '$e');
          }
          cats['$k'] = l;
        }
      });
    }
    for (final g in kGroups) {
      cats.putIfAbsent(g, () => List<String>.from(kDefaultCategories[g]!));
    }
    final imp = <String, ImportedMonth>{};
    if (j['imported'] is Map) {
      (j['imported'] as Map).forEach((k, v) {
        if (v is Map) {
          final m = ImportedMonth.fromJson(Map<String, dynamic>.from(v));
          imp[m.key] = m;
        }
      });
    }
    return MoneyData(
      phoneId: '${j['phoneId'] ?? newId('ph')}',
      phoneName: '${j['phoneName'] ?? ''}',
      role: '${j['role'] ?? ''}',
      places: l('places', Place.fromJson),
      people: l('people', Person.fromJson),
      bills: l('bills', Bill.fromJson),
      txns: l('txns', Txn.fromJson),
      categories: cats,
      imported: imp,
      copySeen: (j['copySeen'] is Map) ? (j['copySeen'] as Map).map((k, v) => MapEntry('$k', '$v')) : null,
      lastPlaceId: '${j['lastPlaceId'] ?? ''}',
      lastInPlaceId: '${j['lastInPlaceId'] ?? ''}',
      sources: sl('sources'),
      reasons: sl('reasons'),
      billNames: j['billNames'] is List ? sl('billNames') : null,
      peerName: '${j['peerName'] ?? ''}',
      syncFolder: '${j['syncFolder'] ?? ''}',
    );
  }

  // ---------- الاختيارات اللي بتتكتب مرة وتفضل ----------
  /// بند جديد تحت النوع ده (من غير تكرار)
  String addCategory(String group, String raw) => addChoice(categories.putIfAbsent(group, () => []), raw);
  String addSource(String raw) => addChoice(sources, raw);
  String addReason(String raw) => addChoice(reasons, raw);
  String addBillName(String raw) => addChoice(billNames, raw);

  /// شخص بالاسم: لو موجود (نفس الاسم بعد التنضيف) بيرجّعه، لو لأ بيضيفه
  Person? findOrAddPerson(String raw, {bool family = false, bool child = false}) {
    final c = cleanLabel(raw);
    if (c.isEmpty) return null;
    for (final p in people) {
      if (sameLabel(p.name, c)) return p;
    }
    final p = Person(id: newId('pe'), name: c, family: family, child: child);
    people.add(p);
    return p;
  }

  // ---------- أسماء ----------
  Place? placeById(String id) {
    for (final p in places) {
      if (p.id == id) return p;
    }
    return null;
  }

  Person? personById(String id) {
    for (final p in people) {
      if (p.id == id) return p;
    }
    return null;
  }

  Bill? billById(String id) {
    for (final b in bills) {
      if (b.id == id) return b;
    }
    return null;
  }

  String placeName(String id) => placeById(id)?.name ?? (id.isEmpty ? '—' : 'مكان اتمسح');
  String personName(String id) => personById(id)?.name ?? (id.isEmpty ? '' : 'شخص اتمسح');

  List<Place> get activePlaces => places.where((p) => !p.archived).toList();

  // ---------- الأرصدة ----------
  /// رصيد المكان = رصيد البداية + اللي دخل − اللي خرج (كل الأوقات).
  /// [upTo] لو موجود: لحد اليوم ده (شامل).
  int balanceOf(String placeId, {String? upTo, String? excludeTxnId}) {
    final p = placeById(placeId);
    var b = p?.opening ?? 0;
    for (final t in txns) {
      if (excludeTxnId != null && t.id == excludeTxnId) continue;
      if (upTo != null && t.date.compareTo(upTo) > 0) continue;
      b += deltaFor(t, placeId);
    }
    return b;
  }

  static int deltaFor(Txn t, String placeId) {
    var d = 0;
    if (t.placeId == placeId) {
      if (isOut(t.type)) {
        d -= t.amount;
      } else {
        d += t.amount; // add, repay
      }
    }
    if (t.type == tTransfer && t.toPlaceId == placeId) d += t.amount;
    return d;
  }

  int get totalBalance =>
      places.fold(0, (s, p) => s + balanceOf(p.id));

  /// الرصيد بعد العملية دي (لو اتسجلت/اتعدلت) — للتحذير من السالب
  int balanceAfter(Txn t, String placeId, {String? replacingId}) =>
      balanceOf(placeId, excludeTxnId: replacingId) + deltaFor(t, placeId);

  /// الأماكن اللي هتبقى بالسالب بعد العملية دي (تحذير بس، مش منع)
  List<String> negativeWarnings(Txn t, {String? replacingId}) {
    final out = <String>[];
    final ids = {t.placeId, if (t.type == tTransfer) t.toPlaceId}..remove('');
    for (final id in ids) {
      if (deltaFor(t, id) >= 0) continue;
      final after = balanceAfter(t, id, replacingId: replacingId);
      if (after < 0) out.add('${placeName(id)} هيبقى ${egp(after)}');
    }
    return out;
  }

  // ---------- العمليات ----------
  /// تسجيل أو تعديل (نفس الـ id = تعديل)
  void upsert(Txn t) {
    final i = txns.indexWhere((e) => e.id == t.id);
    if (i >= 0) {
      txns[i] = t;
    } else {
      txns.add(t);
    }
    if (t.placeId.isNotEmpty) {
      if (isOut(t.type)) {
        lastPlaceId = t.placeId;
      } else {
        lastInPlaceId = t.placeId;
      }
    }
  }

  bool removeTxn(String id) {
    final n = txns.length;
    txns.removeWhere((e) => e.id == id);
    return txns.length != n;
  }

  List<Txn> monthTxns(String mk) =>
      txns.where((t) => inMonth(t.date, mk)).toList()
        ..sort((a, b) {
          final c = b.date.compareTo(a.date);
          return c != 0 ? c : b.createdAt.compareTo(a.createdAt);
        });

  // ---------- السلف ----------
  int advanceRemaining(String personId) {
    var r = 0;
    for (final t in txns) {
      if (t.personId != personId) continue;
      if (t.type == tAdvance) r += t.amount;
      if (t.type == tRepay) r -= t.amount;
    }
    return r;
  }

  /// كل الناس اللي عليهم أو كان عليهم سلف
  List<Person> get advancePeople {
    final ids = txns
        .where((t) => t.type == tAdvance || t.type == tRepay)
        .map((t) => t.personId)
        .toSet();
    return people.where((p) => ids.contains(p.id)).toList();
  }

  int get totalAdvancesOut =>
      advancePeople.fold(0, (s, p) => s + max(0, advanceRemaining(p.id)));

  // ---------- الفواتير ----------
  Txn? billPayment(String billId, String mk) {
    for (final t in txns) {
      if (t.type == tBill && t.billId == billId && t.billMonth == mk) return t;
    }
    return null;
  }

  /// دفع فاتورة شهر: بيعمل مصروف «فاتورة» من المكان المختار
  Txn payBill(Bill b, String mk,
      {required int amount, required String placeId, String? date, String note = ''}) {
    final old = billPayment(b.id, mk);
    final t = Txn(
      id: old?.id ?? newId('tx'),
      type: tBill,
      amount: amount,
      date: date ?? isoDate(DateTime.now()),
      placeId: placeId,
      group: kGroupHome,
      category: b.name,
      billId: b.id,
      billMonth: mk,
      note: note,
    );
    upsert(t);
    return t;
  }

  /// فواتير الشهر اللي لسه متدفعتش ومعادها قرب (خلال [days] أيام) أو عدّى
  List<Bill> dueBills(String mk, DateTime today, {int days = 3}) {
    final out = <Bill>[];
    for (final b in bills.where((b) => b.active)) {
      if (billPayment(b.id, mk) != null) continue;
      final due = b.dueIn(mk);
      final diff = due.difference(DateTime(today.year, today.month, today.day)).inDays;
      if (diff <= days) out.add(b);
    }
    return out;
  }

  // ---------- الملخص ----------
  MonthSummary ownSummary(String mk) =>
      summarize(txns, mk, personName: personName, placeName: placeName);

  MonthSummary importedSummary(ImportedMonth m) => summarize(m.txns, m.month,
      personName: (id) => m.personNames[id] ?? (id.isEmpty ? '' : 'شخص'),
      placeName: (id) => m.placeNames[id] ?? 'مكان');

  List<ImportedMonth> importedFor(String mk) =>
      imported.values.where((m) => m.month == mk).toList()
        ..sort((a, b) => a.phoneName.compareTo(b.phoneName));

  MonthSummary combinedSummary(String mk) {
    final s = ownSummary(mk);
    for (final m in importedFor(mk)) {
      s.addAll(importedSummary(m));
    }
    return s;
  }

  /// أول مرة: مفيش أرصدة ولا عمليات ← «ابدأ بأرصدتك»
  bool get isEmptyStart => txns.isEmpty && places.every((p) => p.opening == 0);

  /// أرصدة البداية لكل الأماكن مرة واحدة (null = سيبه زي ما هو)
  void setOpenings(Map<String, int?> byPlaceId) {
    byPlaceId.forEach((id, v) {
      final p = placeById(id);
      if (p != null && v != null) p.opening = v;
    });
  }

  // ---------- تحويل للتليفون التاني ----------
  /// ملف التحويل (صغير) اللي بيتبعت للتليفون التاني
  Map<String, dynamic> transferPayload(Txn t) => {
        'app': kAppName,
        'type': 'transfer',
        'format': 1,
        'id': t.id,
        'amount': t.amount,
        'date': t.date,
        'note': t.note,
        'fromPhoneId': phoneId,
        'fromPhoneName': badgeName,
        'toPhoneName': peerName,
        'createdAt': t.createdAt,
      };

  bool hasReceivedTransfer(String transferId) =>
      transferId.isNotEmpty && txns.any((t) => t.type == tFromPhone && t.refId == transferId);

  /// استلام ملف تحويل: بيعمل «استلام» في المكان المختار. نفس الملف مرتين = مرفوض. (الاتجاهين)
  TransferResult importTransferJson(dynamic raw, String placeId) {
    if (raw is! Map) return const TransferResult(false, 'الملف ده مش ملف تحويل من A2R Money');
    final j = Map<String, dynamic>.from(raw);
    if (j['app'] != kAppName || j['type'] != 'transfer') {
      return const TransferResult(false, 'الملف ده مش ملف تحويل من A2R Money');
    }
    final id = '${j['id'] ?? ''}';
    final amt = (j['amount'] as num?)?.toInt() ?? 0;
    final date = '${j['date'] ?? ''}';
    if (id.isEmpty || amt <= 0 || !RegExp(r'^\d{4}-\d{2}-\d{2}$').hasMatch(date)) {
      return const TransferResult(false, 'ملف التحويل ناقص أو بايظ');
    }
    if ('${j['fromPhoneId'] ?? ''}' == phoneId) {
      return const TransferResult(false, 'التحويل ده طالع من التليفون ده نفسه — افتحه على التليفون التاني');
    }
    if (hasReceivedTransfer(id)) return const TransferResult(false, 'التحويل ده اتسجل قبل كده');
    if (placeById(placeId) == null) return const TransferResult(false, 'اختار المكان اللي الفلوس هتدخله');
    final note = '${j['note'] ?? ''}'.trim();
    final t = Txn(
      id: newId('tx'),
      type: tFromPhone,
      amount: amt,
      date: date,
      placeId: placeId,
      note: note,
      refId: id,
    );
    upsert(t);
    learnPeer(j['fromPhoneName']);
    return TransferResult(true, 'اتسجل ${egp(amt)} من «${j['fromPhoneName'] ?? otherName}» في ${placeName(placeId)}', txn: t);
  }

  /// مقارنة التحويلات في شهر (الاتجاهين): اللي حوّلته ← اللي التاني استلمه، واللي التاني حوّله ← اللي انا استلمته
  ({int sent, int received, int otherSent, int ownReceived, bool hasOther}) transferCheck(String mk) {
    final own = ownSummary(mk);
    var rec = 0, os = 0;
    final others = importedFor(mk);
    for (final m in others) {
      final x = importedSummary(m);
      rec += x.fromPhone;
      os += x.toPhone;
    }
    return (sent: own.toPhone, received: rec, otherSent: os, ownReceived: own.fromPhone, hasOther: others.isNotEmpty);
  }

  // ---------- ملف القوائم (الأسرة، البنود، الفواتير…) — من غير أي فلوس ----------
  /// القوائم بس: أسماء الناس، البنود، الفواتير (كقالب)، المصادر، الأسباب، أسماء الأماكن.
  /// مفيش فيه: عمليات، أرصدة، رصيد بداية، سلف، دور التليفون.
  Map<String, dynamic> exportLists() => {
        'app': kAppName,
        'type': 'lists',
        'format': 1,
        'phoneId': phoneId,
        'phoneName': badgeName,
        'exportedAt': DateTime.now().toIso8601String(),
        'people': [
          for (final p in people) {'name': p.name, 'family': p.family, 'child': p.child}
        ],
        'categories': {for (final e in categories.entries) e.key: List<String>.from(e.value)},
        'bills': [
          for (final b in bills)
            {'name': b.name, 'kind': b.kind, 'dueDay': b.dueDay, 'usual': b.usual, 'active': b.active}
        ],
        'billNames': List<String>.from(billNames),
        'sources': List<String>.from(sources),
        'reasons': List<String>.from(reasons),
        'places': [
          for (final p in places.where((p) => !p.archived)) {'name': p.name, 'kind': p.kind}
        ],
      };

  /// دمج ملف قوائم: بيضيف الناقص بس (نفس مطابقة الأسماء: همزات/ة-ه/ى-ي/مسافات)،
  /// عمره ما بيمسح أو يغيّر اسم موجود، ومش بيلمس الأرصدة ولا العمليات.
  /// [apply] = false ← حساب بس (للمعاينة) من غير أي تغيير.
  ListsResult mergeListsJson(dynamic raw, {bool addPlaces = true, bool apply = true}) {
    if (raw is! Map) return const ListsResult(false, 'الملف ده مش ملف قوائم من A2R Money');
    final j = Map<String, dynamic>.from(raw);
    if (j['app'] != kAppName || j['type'] != 'lists') {
      return const ListsResult(false, 'الملف ده مش ملف قوائم من A2R Money');
    }
    final pid = '${j['phoneId'] ?? ''}';
    if (pid.isEmpty) return const ListsResult(false, 'ملف القوائم ناقص أو بايظ');
    if (pid == phoneId) {
      return const ListsResult(false, 'الملف ده طالع من التليفون ده نفسه — افتحه على التليفون التاني');
    }
    // المعاينة بتشتغل على نسخة، عشان الحساب يبقى هو هو اللي هيحصل فعلاً
    final t = apply ? this : MoneyData.fromJson(jsonDecode(jsonEncode(toJson())) as Map<String, dynamic>);
    const cap = 1000;
    List<Map<String, dynamic>> maps(String k) => ((j[k] is List) ? j[k] as List : const [])
        .whereType<Map>()
        .take(cap)
        .map((e) => Map<String, dynamic>.from(e))
        .toList();
    List<String> strs(dynamic v) =>
        ((v is List) ? v : const []).take(cap).map((e) => cleanLabel('$e')).where((e) => e.isNotEmpty && e.length <= 80).toList();
    int addAll(List<String> target, List<String> items) {
      var n = 0;
      for (final x in items) {
        final before = target.length;
        addChoice(target, x);
        if (target.length > before) n++;
      }
      return n;
    }

    var nPeople = 0, nCats = 0, nBills = 0, nPlaces = 0, nOther = 0;
    for (final m in maps('people')) {
      final name = cleanLabel('${m['name'] ?? ''}');
      if (name.isEmpty || name.length > 80) continue;
      if (t.people.any((p) => sameLabel(p.name, name))) continue;
      t.people.add(Person(id: newId('pe'), name: name, family: m['family'] != false, child: m['child'] == true));
      nPeople++;
    }
    if (j['categories'] is Map) {
      (j['categories'] as Map).forEach((k, v) {
        final g = cleanLabel('$k');
        if (g.isEmpty || g.length > 40) return;
        nCats += addAll(t.categories.putIfAbsent(g, () => []), strs(v));
      });
    }
    for (final m in maps('bills')) {
      final name = cleanLabel('${m['name'] ?? ''}');
      if (name.isEmpty || name.length > 80) continue;
      t.addBillName(name);
      if (t.bills.any((b) => sameLabel(b.name, name))) continue;
      final kind = cleanLabel('${m['kind'] ?? ''}');
      t.bills.add(Bill(
        id: newId('bi'),
        name: name,
        kind: kind.isEmpty ? 'بيت' : kind,
        dueDay: ((m['dueDay'] as num?)?.toInt() ?? 1).clamp(1, 31),
        usual: max(0, (m['usual'] as num?)?.toInt() ?? 0),
        active: m['active'] != false,
      ));
      nBills++;
    }
    nOther += addAll(t.billNames, strs(j['billNames']));
    nOther += addAll(t.sources, strs(j['sources']));
    nOther += addAll(t.reasons, strs(j['reasons']));
    final missingPlaces = <Map<String, dynamic>>[];
    for (final m in maps('places')) {
      final name = cleanLabel('${m['name'] ?? ''}');
      if (name.isEmpty || name.length > 60) continue;
      if (t.places.any((p) => sameLabel(p.name, name))) continue;
      if (missingPlaces.any((x) => sameLabel('${x['name']}', name))) continue;
      missingPlaces.add({'name': name, 'kind': m['kind']});
    }
    if (addPlaces) {
      for (final m in missingPlaces) {
        final k = '${m['kind'] ?? ''}';
        t.places.add(Place(
            id: newId('pl'),
            name: '${m['name']}',
            kind: const [kPlaceCash, kPlaceBank, kPlaceWallet, kPlaceOther].contains(k) ? k : kPlaceOther));
        nPlaces++;
      }
    }
    return ListsResult(true, '',
        people: nPeople,
        categories: nCats,
        bills: nBills,
        places: nPlaces,
        missingPlaces: missingPlaces.length,
        other: nOther,
        fromName: '${j['phoneName'] ?? 'التليفون التاني'}');
  }

  // ---------- نسخة كاملة (أي وقت، من أي تليفون للتاني) ----------
  /// شهور التليفون ده اللي فيها عمليات
  List<String> ownMonths() {
    final m = <String>{for (final t in txns) if (t.date.length >= 7) t.date.substring(0, 7)};
    return m.toList()..sort();
  }

  /// نسختي: عملياتي (شهر واحد أو كل الشهور) + القوايم. مفيش فيها أي حاجة جاتلي من تليفون تاني.
  Map<String, dynamic> exportCopy({String? month}) {
    final months = month == null ? ownMonths() : [month];
    return {
      'app': kAppName,
      'type': 'copy',
      'format': 1,
      'phoneId': phoneId,
      'phoneName': badgeName,
      'exportedAt': DateTime.now().toIso8601String(),
      'scope': month ?? 'all',
      'months': [for (final mk in months) exportMonth(mk)],
      'lists': exportLists(),
    };
  }

  /// آخر نسخة استلمتها من تليفون (عشان «تحديث» يعرف لو في أحدث)
  String lastCopyAt(String fromPhoneId) {
    var best = '';
    for (final m in imported.values) {
      if (m.phoneId == fromPhoneId && m.exportedAt.compareTo(best) > 0) best = m.exportedAt;
    }
    final c = copySeen[fromPhoneId] ?? '';
    return c.compareTo(best) > 0 ? c : best;
  }

  /// استلام نسخة: بتبدّل اللي جه قبل كده من نفس التليفون (نفس الشهر — أو كل الشهور لو النسخة كاملة).
  /// أرصدة وعمليات التليفون ده مابتتلمسش. القوايم بيتضاف الناقص منها بس (من غير أماكن).
  CopyResult importCopyJson(dynamic raw) {
    if (raw is! Map) return const CopyResult(false, 'الملف ده مش نسخة من A2R Money');
    final j = Map<String, dynamic>.from(raw);
    if (j['app'] != kAppName) return const CopyResult(false, 'الملف ده مش نسخة من A2R Money');
    if (j['type'] == 'month') {
      final r = importMonthJson(j);
      return CopyResult(r.ok, r.message, months: r.ok ? 1 : 0, fromName: r.month?.phoneName ?? '', exportedAt: r.month?.exportedAt ?? '');
    }
    if (j['type'] != 'copy') return const CopyResult(false, 'الملف ده مش نسخة من A2R Money');
    final pid = '${j['phoneId'] ?? ''}';
    if (pid.isEmpty) return const CopyResult(false, 'النسخة ناقصة أو بايظة');
    if (pid == phoneId) return const CopyResult(false, 'النسخة دي طالعة من التليفون ده نفسه — افتحها على التليفون التاني');
    final name = cleanLabel('${j['phoneName'] ?? ''}');
    final at = '${j['exportedAt'] ?? ''}';
    final scope = '${j['scope'] ?? 'all'}';
    final ms = ((j['months'] is List) ? j['months'] as List : const []).whereType<Map>().toList();
    final good = <Map<String, dynamic>>[];
    for (final m in ms) {
      final mm = Map<String, dynamic>.from(m);
      if ('${mm['phoneId']}' != pid || !RegExp(r'^\d{4}-\d{2}$').hasMatch('${mm['month']}')) continue;
      good.add({...mm, 'app': kAppName, 'type': 'month', 'exportedAt': at, 'phoneName': name.isEmpty ? mm['phoneName'] : name});
    }
    if (scope == 'all') {
      // نسخة كاملة: اللي كان جاي من التليفون ده ومش موجود فيها يتشال (اتمسح عنده)
      imported.removeWhere((k, v) => v.phoneId == pid);
    }
    var n = 0;
    for (final mm in good) {
      if (importMonthJson(mm).ok) n++;
    }
    var listsAdded = 0;
    if (j['lists'] is Map) {
      final lr = mergeListsJson(j['lists'], addPlaces: false);
      if (lr.ok) listsAdded = lr.total;
    }
    copySeen[pid] = at;
    learnPeer(name);
    final who = name.isEmpty ? otherName : name;
    return CopyResult(true, 'اتحدث من «$who» — $n شهر${listsAdded > 0 ? ' • اتضاف $listsAdded من القوايم' : ''}',
        months: n, listsAdded: listsAdded, fromName: who, exportedAt: at);
  }

  /// «تحديث»: من كل الملفات اللي في فولدر الاستلام بيختار الأحدث (بالوقت اللي جوه الملف، مش الاسم)
  SyncPlan planSync(List<Map<String, dynamic>> files) {
    DateTime? when(Map<String, dynamic> j) => DateTime.tryParse('${j['exportedAt'] ?? j['createdAt'] ?? ''}');
    final newestCopy = <String, Map<String, dynamic>>{}; // تليفون ← أحدث نسخة كاملة/شهر
    final newestMonth = <String, Map<String, dynamic>>{}; // تليفون|شهر ← أحدث ملف شهر قديم (v1/v2)
    final newestLists = <String, Map<String, dynamic>>{};
    final transfers = <String, Map<String, dynamic>>{};
    bool newer(Map<String, dynamic> a, Map<String, dynamic>? b) {
      if (b == null) return true;
      final wa = when(a), wb = when(b);
      if (wa == null) return false;
      if (wb == null) return true;
      return wa.isAfter(wb);
    }

    for (final j in files) {
      if (j['app'] != kAppName) continue;
      final t = '${j['type']}';
      if (t == 'transfer') {
        if ('${j['fromPhoneId']}' == phoneId) continue;
        final id = '${j['id'] ?? ''}';
        if (id.isEmpty || hasReceivedTransfer(id)) continue;
        transfers[id] = j;
        continue;
      }
      final pid = '${j['phoneId'] ?? ''}';
      if (pid.isEmpty || pid == phoneId || when(j) == null) continue;
      if (t == 'copy') {
        if (newer(j, newestCopy[pid])) newestCopy[pid] = j;
      } else if (t == 'month') {
        final k = '$pid|${j['month']}';
        if (newer(j, newestMonth[k])) newestMonth[k] = j;
      } else if (t == 'lists') {
        if (newer(j, newestLists[pid])) newestLists[pid] = j;
      }
    }
    final copies = <Map<String, dynamic>>[];
    newestCopy.forEach((pid, j) {
      final seen = DateTime.tryParse(lastCopyAt(pid));
      final w = when(j)!;
      if (seen == null || w.isAfter(seen)) copies.add(j);
    });
    final months = <Map<String, dynamic>>[];
    newestMonth.forEach((k, j) {
      final pid = '${j['phoneId']}';
      final w = when(j)!;
      // لو في نسخة كاملة أحدث منه هتغطيه
      final c = newestCopy[pid];
      if (c != null && !when(c)!.isBefore(w)) return;
      final old = imported[k];
      final seen = DateTime.tryParse(old?.exportedAt ?? '');
      final cs = DateTime.tryParse(copySeen[pid] ?? '');
      if ((seen == null || w.isAfter(seen)) && (cs == null || w.isAfter(cs))) months.add(j);
    });
    final lists = <Map<String, dynamic>>[];
    newestLists.forEach((pid, j) {
      if (mergeListsJson(j, apply: false, addPlaces: false).total > 0) lists.add(j);
    });
    final tr = transfers.values.toList()..sort((a, b) => '${a['date']}'.compareTo('${b['date']}'));
    return SyncPlan(copies: copies, months: months, lists: lists, transfers: tr);
  }

  // ---------- ملف الشهر (من التليفون التاني) ----------
  Map<String, dynamic> exportMonth(String mk) {
    final ts = monthTxns(mk);
    final last = isoDate(monthLast(mk));
    return {
      'app': kAppName,
      'type': 'month',
      'format': 1,
      'phoneId': phoneId,
      'phoneName': badgeName,
      'month': mk,
      'exportedAt': DateTime.now().toIso8601String(),
      'txns': ts.map((e) => e.toJson()).toList(),
      'placeNames': {for (final p in places) p.id: p.name},
      'personNames': {for (final p in people) p.id: p.name},
      'billNames': {for (final b in bills) b.id: b.name},
      'endBalances': {for (final p in places) p.name: balanceOf(p.id, upTo: last)},
    };
  }

  /// استيراد ملف شهر: لو نفس التليفون ونفس الشهر اتضاف قبل كده ← بيتبدل (مفيش تكرار).
  /// مش بيلمس أرصدة التليفون ده خالص.
  ImportResult importMonthJson(dynamic raw) {
    if (raw is! Map) return const ImportResult(false, 'الملف ده مش ملف شهر من A2R Money');
    final j = Map<String, dynamic>.from(raw);
    if (j['app'] != kAppName || j['type'] != 'month') {
      return const ImportResult(false, 'الملف ده مش ملف شهر من A2R Money');
    }
    final mk = '${j['month'] ?? ''}';
    if (!RegExp(r'^\d{4}-\d{2}$').hasMatch(mk)) {
      return const ImportResult(false, 'الملف مافيهوش شهر صحيح');
    }
    final pid = '${j['phoneId'] ?? ''}';
    if (pid.isEmpty) return const ImportResult(false, 'الملف ناقص (مافيهوش رقم التليفون)');
    if (pid == phoneId) {
      return const ImportResult(false, 'الملف ده طالع من التليفون ده نفسه — افتحه على التليفون التاني');
    }
    final m = ImportedMonth.fromJson(j);
    // نتأكد إن العمليات كلها في نفس الشهر، ومفيش id متكرر جوه الملف
    final seen = <String>{};
    final clean = <Txn>[];
    for (final t in m.txns) {
      if (!inMonth(t.date, mk)) continue;
      if (!seen.add(t.id)) continue;
      clean.add(t);
    }
    final fixed = ImportedMonth(
      phoneId: m.phoneId,
      phoneName: m.phoneName,
      month: m.month,
      exportedAt: m.exportedAt,
      txns: clean,
      placeNames: m.placeNames,
      personNames: m.personNames,
      billNames: m.billNames,
      endBalances: m.endBalances,
    );
    final replaced = imported.containsKey(fixed.key);
    imported[fixed.key] = fixed;
    learnPeer(fixed.phoneName);
    return ImportResult(
        true,
        replaced
            ? 'اتحدّث شهر ${monthLabel(mk)} من «${fixed.phoneName}» (اتبدل القديم، مفيش تكرار)'
            : 'اتضاف شهر ${monthLabel(mk)} من «${fixed.phoneName}»',
        replaced: replaced,
        month: fixed);
  }
}

// ==================== Excel (نفس تنضيف A2R Team عشان يفتح بسرعة في WPS) ====================

String colIndexToLetters(int col1Based) {
  int n = col1Based;
  String s = '';
  while (n > 0) {
    int rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = (n - 1) ~/ 26;
  }
  return s;
}

int colLettersToIndex(String cellRef) {
  final letters = cellRef.replaceAll(RegExp(r'[0-9]'), '');
  int res = 0;
  for (int i = 0; i < letters.length; i++) {
    res = res * 26 + (letters.codeUnitAt(i) - 64);
  }
  return res;
}
List<int> tidyExcelPackageXlsx(List<int> bytes) {
  try {
    final input = ZipDecoder().decodeBytes(bytes);
    final files = <String, List<int>>{};
    for (final f in input.files) {
      if (f.isFile) files[f.name] = f.content as List<int>;
    }
    String? txt(String n) => files[n] == null ? null : utf8.decode(files[n]!);
    void put(String n, String v) => files[n] = utf8.encode(v);

    var ct = txt('[Content_Types].xml') ?? '';
    // 1) الرسم الفاضي
    final emptyDrawings = <String>[];
    for (final n in files.keys.toList()) {
      if (!n.startsWith('xl/drawings/') || !n.endsWith('.xml')) continue;
      final d = txt(n)!;
      if (RegExp(r'<xdr:wsDr[^>]*/>\s*$').hasMatch(d.trim())) {
        emptyDrawings.add(n.substring('xl/'.length));
      }
    }
    for (final n in files.keys.toList()) {
      if (!RegExp(r'^xl/worksheets/sheet\d+\.xml$').hasMatch(n)) continue;
      var sheet = txt(n)!;
      final relName =
          'xl/worksheets/_rels/${n.split('/').last}.rels';
      final rels = txt(relName);
      if (rels != null && emptyDrawings.isNotEmpty) {
        var newRels = rels;
        for (final m in RegExp(r'<Relationship [^>]*?/>').allMatches(rels)) {
          final rel = m.group(0)!;
          final tgt = RegExp(r'Target="\.\./([^"]+)"').firstMatch(rel)?.group(1);
          final id = RegExp(r'Id="([^"]+)"').firstMatch(rel)?.group(1);
          if (tgt != null && id != null && emptyDrawings.contains(tgt)) {
            newRels = newRels.replaceFirst(rel, '');
            sheet = sheet.replaceAll(RegExp('<drawing r:id="$id"\\s*/>'), '');
          }
        }
        if (!newRels.contains('<Relationship ')) {
          files.remove(relName);
        } else {
          put(relName, newRels);
        }
      }
      // 2) حجم الجدول
      if (!sheet.contains('<dimension')) {
        var maxR = 1, maxC = 1;
        for (final m in RegExp(r'<c r="([A-Z]+)(\d+)"').allMatches(sheet)) {
          final c = colLettersToIndex(m.group(1)!);
          final r = int.parse(m.group(2)!);
          if (c > maxC) maxC = c;
          if (r > maxR) maxR = r;
        }
        final dim = '<dimension ref="A1:${colIndexToLetters(maxC)}$maxR"/>';
        if (sheet.contains('</sheetPr>')) {
          sheet = sheet.replaceFirst('</sheetPr>', '</sheetPr>$dim');
        } else {
          sheet = sheet.replaceFirst('<sheetViews>', '$dim<sheetViews>');
        }
      }
      put(n, sheet);
    }
    for (final d in emptyDrawings) {
      files.remove('xl/$d');
      ct = ct.replaceAll(
          RegExp('<Override [^>]*PartName="/xl/$d"[^>]*/>'), '');
    }
    // 3) أبعاد نافذة الملف (bookViews) زي ملفات Excel العادية
    final wb = txt('xl/workbook.xml');
    if (wb != null && !wb.contains('<bookViews')) {
      put('xl/workbook.xml',
          wb.replaceFirst('<sheets>', '<bookViews><workbookView activeTab="0"/></bookViews><sheets>'));
    }
    // \r بعد سطر <?xml ...?> في ورقة الشيت (من مكتبة excel) ← سطر عادي.
    // ملفات Excel الأصلية (القوالب، فيها docProps) بتفضل زي ما هي بالظبط.
    final fromExcelPackage = !files.containsKey('docProps/app.xml');
    for (final n in fromExcelPackage ? files.keys.toList() : const <String>[]) {
      if (!RegExp(r'^xl/worksheets/sheet\d+\.xml$').hasMatch(n)) continue;
      final t = txt(n)!;
      final k = t.indexOf('?>\r');
      if (k >= 0 && k < 100 && !t.startsWith('?>\r\n', k)) {
        put(n, t.replaceFirst('?>\r', '?>\n'));
      }
    }
    // 4) بيانات الملف
    if (!files.containsKey('docProps/app.xml')) {
      final now = DateTime.now().toUtc().toIso8601String().split('.').first;
      put('docProps/app.xml',
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
          'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
          '<Application>Microsoft Excel</Application></Properties>');
      put('docProps/core.xml',
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
          'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
          'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
          '<dc:creator>A2R Money</dc:creator>'
          '<dcterms:created xsi:type="dcterms:W3CDTF">${now}Z</dcterms:created>'
          '<dcterms:modified xsi:type="dcterms:W3CDTF">${now}Z</dcterms:modified>'
          '</cp:coreProperties>');
      ct = ct.replaceFirst(
          '</Types>',
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
              '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
      final root = txt('_rels/.rels');
      if (root != null && !root.contains('docProps/core.xml')) {
        put(
            '_rels/.rels',
            root.replaceFirst(
                '</Relationships>',
                '<Relationship Id="rId90" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                    '<Relationship Id="rId91" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'));
      }
    }
    put('[Content_Types].xml', ct);
    final out = Archive();
    // [Content_Types].xml الأول زي ملفات Excel العادية
    final order = files.keys.toList()
      ..sort((a, b) => a == '[Content_Types].xml'
          ? -1
          : b == '[Content_Types].xml'
              ? 1
              : 0);
    for (final n in order) {
      out.addFile(ArchiveFile(n, files[n]!.length, files[n]!));
    }
    return ZipEncoder().encode(out) ?? bytes;
  } catch (_) {
    return bytes; // أي مشكلة: الملف الأصلي زي ما هو
  }
}

/// شيت الشهر: ملخص + كل العمليات (التليفونين) + الأسرة + السلف
List<int>? buildMonthExcel(MoneyData d, String mk) {
  final ex = xls.Excel.createExcel();
  final def = ex.getDefaultSheet() ?? 'Sheet1';
  xls.CellValue tv(Object? v) => v is int
      ? xls.IntCellValue(v)
      : v is double
          ? xls.DoubleCellValue(v)
          : xls.TextCellValue('${v ?? ''}');
  double pounds(int p) => p / 100.0;
  void rows(String name, List<List<Object?>> data) {
    final sh = ex[name];
    sh.isRTL = true;
    for (final r in data) {
      sh.appendRow(r.map(tv).toList());
    }
    for (var c = 0; c < (data.isEmpty ? 0 : data.first.length); c++) {
      sh.setColumnWidth(c, c == 0 ? 30.0 : 16.0);
    }
  }

  final own = d.ownSummary(mk);
  final ownName = d.phoneName.isEmpty ? 'التليفون ده' : d.phoneName;
  final all = d.combinedSummary(mk);
  final sum = <List<Object?>>[
    ['A2R Money — ملخص ${monthLabel(mk)}', ''],
    ['', ''],
    ['البند', 'الإجمالي (جنيه)'],
    ['إجمالي الصرف (مصروفات + فواتير + سحب)', pounds(all.totalSpend)],
    ['المصروفات', pounds(all.expenses)],
    ['الفواتير', pounds(all.bills)],
    ['السحب', pounds(all.withdraws)],
    ['الفلوس اللي اتضافت', pounds(all.added)],
    ['سلف اتدفعت', pounds(all.advancesGiven)],
    ['سلف رجعت', pounds(all.advancesBack)],
    ['اتحول لتليفون تاني (مش صرف)', pounds(own.toPhone)],
    ['اتاستلم من تليفون تاني (مش صرف)', pounds(all.fromPhone)],
    ['', ''],
    ['حسب التليفون', ''],
    [ownName, pounds(own.totalSpend)],
    for (final m in d.importedFor(mk)) [m.phoneName, pounds(d.importedSummary(m).totalSpend)],
    ['', ''],
    ['حسب النوع', ''],
    for (final e in (all.byGroup.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
      [e.key, pounds(e.value)],
    ['', ''],
    ['حسب البند', ''],
    for (final e in (all.byCategory.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
      [e.key, pounds(e.value)],
  ];
  rows('الملخص', sum);

  final ops = <List<Object?>>[
    ['التليفون', 'التاريخ', 'النوع', 'المبلغ (جنيه)', 'المكان', 'لـ مكان', 'المجموعة', 'البند', 'الشخص', 'ملاحظة'],
  ];
  void addOps(String phone, List<Txn> ts, String Function(String) pl, String Function(String) pe) {
    for (final t in ts) {
      ops.add([
        phone,
        t.date,
        kTypeLabels[t.type] ?? t.type,
        pounds(t.amount),
        pl(t.placeId),
        t.type == tTransfer ? pl(t.toPlaceId) : '',
        t.group,
        t.category,
        t.personId.isEmpty ? '' : pe(t.personId),
        t.note,
      ]);
    }
  }

  addOps(ownName, d.monthTxns(mk), d.placeName, d.personName);
  for (final m in d.importedFor(mk)) {
    final ts = List<Txn>.from(m.txns)..sort((a, b) => b.date.compareTo(a.date));
    addOps(m.phoneName, ts, (id) => m.placeNames[id] ?? 'مكان', (id) => m.personNames[id] ?? 'شخص');
  }
  rows('العمليات', ops);

  rows('الأسرة', [
    ['الفرد', 'صرف الشهر (جنيه)'],
    for (final e in (all.byPerson.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
      [e.key, pounds(e.value)],
  ]);
  rows('السلف', [
    ['الشخص', 'الباقي عليه (جنيه) — التليفون ده'],
    for (final p in d.advancePeople) [p.name, pounds(d.advanceRemaining(p.id))],
  ]);
  rows('الأرصدة', [
    ['التليفون', 'المكان', 'الرصيد آخر الشهر (جنيه)'],
    for (final p in d.places) [ownName, p.name, pounds(d.balanceOf(p.id, upTo: isoDate(monthLast(mk))))],
    for (final m in d.importedFor(mk))
      for (final e in m.endBalances.entries) [m.phoneName, e.key, pounds(e.value)],
  ]);
  if (def != 'الملخص') ex.delete(def);
  ex.setDefaultSheet('الملخص');
  final bytes = ex.save();
  if (bytes == null) return null;
  return tidyExcelPackageXlsx(bytes);
}

// ==================== الملفات والنسخ الاحتياطي ====================

Future<Directory> getAppFolder({bool requestPermissions = true}) async {
  if (Platform.isAndroid) {
    if (requestPermissions) {
      final ok = await Permission.manageExternalStorage.isGranted;
      if (!ok) {
        await Permission.storage.request();
        await Permission.manageExternalStorage.request();
      }
    }
    final ext = Directory('/storage/emulated/0/$kFolder');
    try {
      if (!await ext.exists()) await ext.create(recursive: true);
      return ext;
    } catch (_) {}
  }
  final appDir = await getApplicationDocumentsDirectory();
  final f = Directory('${appDir.path}/$kFolder');
  if (!await f.exists()) await f.create(recursive: true);
  return f;
}

Future<bool> hasSharedStorageAccess() async {
  if (!Platform.isAndroid) return true;
  try {
    if (await Permission.manageExternalStorage.isGranted) return true;
    if (await Permission.storage.isGranted) {
      final d = Directory('/storage/emulated/0/$kFolder');
      if (!await d.exists()) await d.create(recursive: true);
      final t = File('${d.path}/.write_test');
      await t.writeAsString('ok');
      await t.delete();
      return true;
    }
  } catch (_) {}
  return false;
}

const String kBackupsFolder = 'Backups';
const String kAutoPrefix = 'A2R_Money_Auto_';
const String kBeforeRestorePrefix = 'A2R_Money_BeforeRestore_';
const String kManualBackupName = 'A2R_Money_Backup.json';
const String kMonthFilePrefix = 'A2R_Money_Month_';
const int kKeepAutoBackups = 14;
const int kKeepSafetyBackups = 5;

const List<String> kExtraFolders = [
  '/storage/emulated/0/Download',
  '/storage/emulated/0/Documents',
  '/storage/emulated/0/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Documents',
  '/storage/emulated/0/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Documents/Sent',
  '/storage/emulated/0/Android/media/com.whatsapp.w4b/WhatsApp Business/Media/WhatsApp Business Documents',
  '/storage/emulated/0/WhatsApp/Media/WhatsApp Documents',
];

/// فولدرات مستندات واتساب الموجودة فعلاً على الموبايل — الشكلين:
///  • العادي: Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Documents
///  • الجديد (أكتر من حساب): Android/media/com.whatsapp/WhatsApp/accounts/<رقم>/Media/WhatsApp Documents
/// ونفس الكلام لواتساب بيزنس. الأول في القايمة = الاقتراح الأساسي ([path, الاسم]).
Future<List<List<String>>> discoverWhatsAppFolders({String root = '/storage/emulated/0'}) async {
  final out = <List<String>>[];
  final seen = <String>{};
  Future<void> add(String path, String label) async {
    if (seen.contains(path)) return;
    try {
      if (await Directory(path).exists()) {
        seen.add(path);
        out.add([path, label]);
      }
    } catch (_) {}
  }

  for (final a in const [
    ['com.whatsapp', 'WhatsApp', 'WhatsApp Documents', 'واتساب'],
    ['com.whatsapp.w4b', 'WhatsApp Business', 'WhatsApp Business Documents', 'واتساب بيزنس'],
  ]) {
    final base = '$root/Android/media/${a[0]}/${a[1]}';
    // الشكل الجديد: accounts/<أي رقم>/Media/...
    final accs = <String>[];
    try {
      final ad = Directory('$base/accounts');
      if (await ad.exists()) {
        await for (final e in ad.list()) {
          if (e is Directory) accs.add(e.uri.pathSegments.where((x) => x.isNotEmpty).last);
        }
      }
    } catch (_) {}
    accs.sort();
    for (final id in accs) {
      await add('$base/accounts/$id/Media/${a[2]}', accs.length > 1 ? '${a[3]} — مستندات (حساب $id)' : '${a[3]} — مستندات');
    }
    // الشكل العادي
    await add('$base/Media/${a[2]}', accs.isEmpty ? '${a[3]} — مستندات' : '${a[3]} — مستندات (الفولدر العادي)');
  }
  // موبايلات قديمة (قبل أندرويد 11)
  await add('$root/WhatsApp/Media/WhatsApp Documents', 'واتساب — مستندات (موبايل قديم)');
  await add('$root/WhatsApp Business/Media/WhatsApp Business Documents', 'واتساب بيزنس — مستندات (موبايل قديم)');
  return out;
}

/// كل الفولدرات اللي بندوّر فيها بإيد: القايمة القديمة + فولدرات واتساب اللي اتلقت (بالشكلين) + Sent و Private جواها
Future<List<String>> scanFolders({String root = '/storage/emulated/0'}) async {
  final out = <String>[];
  void add(String p) {
    if (!out.contains(p)) out.add(p);
  }

  for (final w in await discoverWhatsAppFolders(root: root)) {
    add(w[0]);
    add('${w[0]}/Sent');
    add('${w[0]}/Private');
  }
  for (final p in kExtraFolders) {
    add(root == '/storage/emulated/0' ? p : p.replaceFirst('/storage/emulated/0', root));
  }
  return out;
}

class FoundFile {
  final File file;
  final DateTime when;
  final String label;
  final Map<String, dynamic> json;
  FoundFile(this.file, this.when, this.label, this.json);
}

Future<Map<String, dynamic>?> _readJson(File f) async {
  try {
    final raw = jsonDecode(await f.readAsString());
    if (raw is Map) return Map<String, dynamic>.from(raw);
  } catch (_) {}
  return null;
}

/// نسخ احتياطية (التلقائية + اليدوية + اللي اتنقلت لـ Download/واتساب)
Future<List<FoundFile>> listBackups() async {
  final out = <FoundFile>[];
  final seen = <String>{};
  Future<void> consider(File f, String kind) async {
    if (!seen.add(f.path)) return;
    final j = await _readJson(f);
    if (j == null || j['app'] != kAppName || j['places'] is! List || j['type'] != null) return;
    final when = DateTime.tryParse('${j['savedAt'] ?? ''}') ?? await f.lastModified();
    final n = (j['txns'] as List?)?.length ?? 0;
    out.add(FoundFile(f, when, '$kind • ${arDate(when)} ${two(when.hour)}:${two(when.minute)} • $n عملية', j));
  }

  try {
    final dir = await getAppFolder(requestPermissions: false);
    final manual = File('${dir.path}/$kManualBackupName');
    if (await manual.exists()) await consider(manual, 'يدوي');
    final b = Directory('${dir.path}/$kBackupsFolder');
    if (await b.exists()) {
      await for (final e in b.list()) {
        if (e is File && e.path.endsWith('.json')) {
          final n = e.uri.pathSegments.last;
          await consider(e, n.startsWith(kBeforeRestorePrefix) ? 'قبل الاسترجاع' : 'تلقائي');
        }
      }
    }
  } catch (_) {}
  for (final p in await scanFolders()) {
    try {
      final d = Directory(p);
      if (!await d.exists()) continue;
      await for (final e in d.list()) {
        if (e is! File) continue;
        final n = e.uri.pathSegments.last;
        if (!n.endsWith('.json') || !n.startsWith('A2R_Money') || n.startsWith(kMonthFilePrefix)) continue;
        await consider(e, 'ملف ($n)');
      }
    } catch (_) {}
  }
  out.sort((a, b) => b.when.compareTo(a.when));
  return out;
}

/// ملفات الشهر اللي جاية من التليفون التاني (واتساب / Download / فولدر التطبيق)
Future<List<FoundFile>> listMonthFiles(String ownPhoneId) async {
  final out = <FoundFile>[];
  final seen = <String>{};
  final dirs = <String>[...await scanFolders()];
  try {
    final dir = await getAppFolder(requestPermissions: false);
    dirs.insert(0, '${dir.path}/Months');
  } catch (_) {}
  for (final p in dirs) {
    try {
      final d = Directory(p);
      if (!await d.exists()) continue;
      await for (final e in d.list()) {
        if (e is! File || !seen.add(e.path)) continue;
        final n = e.uri.pathSegments.last;
        if (!n.endsWith('.json') || !n.startsWith('A2R_Money')) continue;
        final j = await _readJson(e);
        if (j == null || j['app'] != kAppName || (j['type'] != 'month' && j['type'] != 'copy')) continue;
        if ('${j['phoneId']}' == ownPhoneId) continue;
        final when = DateTime.tryParse('${j['exportedAt'] ?? ''}') ?? await e.lastModified();
        if (j['type'] == 'copy') {
          final ms = (j['months'] as List?) ?? const [];
          var cnt = 0;
          for (final m in ms) {
            if (m is Map) cnt += (m['txns'] as List?)?.length ?? 0;
          }
          out.add(FoundFile(e, when,
              '${j['phoneName']} • نسخة ${j['scope'] == 'all' ? 'كل الشهور' : monthLabel('${j['scope']}')} • $cnt عملية', j));
          continue;
        }
        final mk = '${j['month']}';
        final cnt = (j['txns'] as List?)?.length ?? 0;
        out.add(FoundFile(e, when,
            '${j['phoneName']} • ${RegExp(r'^\d{4}-\d{2}$').hasMatch(mk) ? monthLabel(mk) : mk} • $cnt عملية', j));
      }
    } catch (_) {}
  }
  out.sort((a, b) => b.when.compareTo(a.when));
  return out;
}

const String kTransferFilePrefix = 'A2R_Money_Transfer_';

const String kListsFilePrefix = 'A2R_Money_Lists_';
const String kCopyFilePrefix = 'A2R_Money_Copy_';

/// فولدرات الاستلام المقترحة (أول واحد = واتساب، هو الاقتراح الأساسي)
const List<List<String>> kSyncFolderChoices = [
  ['/storage/emulated/0/Download', 'التنزيلات (Download)'],
  ['/storage/emulated/0/Documents', 'المستندات (Documents)'],
];

/// كل ملفات A2R Money (JSON) اللي في فولدر — بالمحتوى مش بالاسم (واتساب ساعات بيزوّد «(1)» للاسم)
Future<List<Map<String, dynamic>>> readSyncFolder(String path) async {
  final out = <Map<String, dynamic>>[];
  final d = Directory(path);
  if (!await d.exists()) return out;
  // الفولدر نفسه + Private و Sent جواه (واتساب الجديد ساعات بيحط الملفات فيهم)
  final ents = <FileSystemEntity>[];
  for (final dd in [d, Directory('$path/Private'), Directory('$path/Sent')]) {
    try {
      if (await dd.exists()) ents.addAll(await dd.list().toList());
    } catch (_) {}
  }
  for (final e in ents) {
    if (e is! File) continue;
    final n = e.uri.pathSegments.last.toLowerCase();
    if (!n.endsWith('.json')) continue;
    try {
      if (await e.length() > 8 * 1024 * 1024) continue;
    } catch (_) {
      continue;
    }
    final j = await _readJson(e);
    if (j != null && j['app'] == kAppName && j['type'] != null) out.add(j);
  }
  return out;
}

/// عدد ملفات A2R Money في فولدر (للاختيار)
Future<int> countSyncFiles(String path) async {
  try {
    final d = Directory(path);
    if (!await d.exists()) return -1;
    var n = 0;
    await for (final e in d.list()) {
      if (e is File && e.uri.pathSegments.last.startsWith('A2R_Money') && e.path.endsWith('.json')) n++;
    }
    return n;
  } catch (_) {
    return -1;
  }
}

/// اختيار فولدر الاستلام (مرة واحدة، ويتغير من الإعدادات)
class FolderPickerScreen extends StatefulWidget {
  final String current;
  final String root; // جذر التخزين (للاختبار بس بيتغير)
  const FolderPickerScreen({super.key, this.current = '', this.root = '/storage/emulated/0'});
  @override
  State<FolderPickerScreen> createState() => _FolderPickerScreenState();
}

class _FolderPickerScreenState extends State<FolderPickerScreen> {
  Map<String, int> counts = {};
  List<List<String>> choices = []; // فولدرات واتساب اللي اتلقت الأول، وبعدين Download/Documents
  bool loaded = false;
  String? browse; // لو بيتصفح: الفولدر الحالي
  List<Directory> subs = [];

  @override
  void initState() {
    super.initState();
    () async {
      final r = widget.root;
      String fix(String x) => x.replaceFirst('/storage/emulated/0', r);
      final wa = await discoverWhatsAppFolders(root: r);
      final list = <List<String>>[
        ...wa,
        if (wa.isEmpty) [fix('/storage/emulated/0/Android/media/com.whatsapp/WhatsApp/Media/WhatsApp Documents'), 'واتساب — مستندات'],
        for (final c in kSyncFolderChoices) [fix(c[0]), c[1]],
      ];
      final m = <String, int>{};
      for (final c in list) {
        m[c[0]] = await countSyncFiles(c[0]);
      }
      if (mounted) {
        setState(() {
          choices = list;
          counts = m;
          loaded = true;
        });
      }
    }();
  }

  Future<void> open(String path) async {
    final l = <Directory>[];
    try {
      await for (final e in Directory(path).list()) {
        if (e is Directory && !e.uri.pathSegments.where((x) => x.isNotEmpty).last.startsWith('.')) l.add(e);
      }
    } catch (_) {}
    l.sort((a, b) => a.path.toLowerCase().compareTo(b.path.toLowerCase()));
    setState(() {
      browse = path;
      subs = l;
    });
  }

  String nameOf(String path) => path.split('/').where((x) => x.isNotEmpty).last;

  @override
  Widget build(BuildContext context) {
    if (browse != null) {
      final b = browse!;
      final atRoot = b == '/storage/emulated/0';
      return Scaffold(
        appBar: AppBar(title: Text(atRoot ? 'ذاكرة الموبايل' : nameOf(b))),
        body: ListView(padding: const EdgeInsets.fromLTRB(12, 4, 12, 100), children: [
          Text(b, style: const TextStyle(color: C.muted, fontSize: 11.5), textDirection: TextDirection.ltr),
          const SizedBox(height: 6),
          if (!atRoot)
            ListTile(
              leading: const Icon(Icons.arrow_upward_rounded),
              title: const Text('رجوع لفوق'),
              onTap: () => open(b.substring(0, b.lastIndexOf('/'))),
            ),
          if (subs.isEmpty) const Padding(padding: EdgeInsets.all(16), child: Text('مفيش فولدرات جوه ده', style: TextStyle(color: C.muted))),
          for (final d in subs)
            ListTile(
              leading: const Icon(Icons.folder_rounded, color: C.gold),
              title: Text(nameOf(d.path), textDirection: TextDirection.ltr, textAlign: TextAlign.start),
              onTap: () => open(d.path),
            ),
        ]),
        bottomNavigationBar: SafeArea(
          child: Padding(
            padding: const EdgeInsets.fromLTRB(16, 6, 16, 12),
            child: FilledButton.icon(
              style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
              onPressed: () => Navigator.pop(context, b),
              icon: const Icon(Icons.check_rounded),
              label: const Text('اختار الفولدر ده'),
            ),
          ),
        ),
      );
    }
    return Scaffold(
      appBar: AppBar(title: const Text('فولدر الاستلام')),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 4, 16, 40), children: [
        const Text('اختار الفولدر اللي بتنزل فيه الملفات اللي جاية من التليفون التاني. بيتحفظ مرة واحدة، '
            'وبعد كده زرار «تحديث» بياخد أحدث نسخة منه لوحده. (لازم تفتح الملف في شات واتساب مرة عشان يتنزل)',
            style: TextStyle(color: C.muted, fontSize: 12.5)),
        const SizedBox(height: 10),
        if (!loaded) const Padding(padding: EdgeInsets.all(20), child: Center(child: CircularProgressIndicator())),
        for (final c in choices)
          Padding(
            padding: const EdgeInsets.only(bottom: 8),
            child: ListTile(
              tileColor: C.card,
              shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(14),
                  side: BorderSide(color: widget.current == c[0] ? C.gold : C.line, width: widget.current == c[0] ? 1.6 : 1)),
              leading: Icon(c[1].startsWith('واتساب') ? Icons.chat_rounded : Icons.folder_rounded,
                  color: (counts[c[0]] ?? -1) < 0 ? C.muted : C.emerald),
              title: Text(c[1] + (c == choices.first ? ' (مقترح)' : ''), style: const TextStyle(fontWeight: FontWeight.w700)),
              subtitle: Text(
                  !counts.containsKey(c[0])
                      ? '…'
                      : counts[c[0]]! < 0
                          ? 'مش موجود على الموبايل ده'
                          : counts[c[0]] == 0
                              ? 'موجود • لسه مفيش ملفات A2R Money'
                              : 'موجود • فيه ${counts[c[0]]} ملف A2R Money',
                  style: const TextStyle(color: C.muted, fontSize: 12)),
              trailing: widget.current == c[0] ? const Icon(Icons.check_circle_rounded, color: C.gold) : null,
              onTap: () => Navigator.pop(context, c[0]),
            ),
          ),
        const SizedBox(height: 6),
        OutlinedButton.icon(
          style: OutlinedButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
          onPressed: () => open('/storage/emulated/0'),
          icon: const Icon(Icons.folder_open_rounded),
          label: const Text('فولدر تاني… (تصفّح)'),
        ),
        if (widget.current.isNotEmpty && !choices.any((c) => c[0] == widget.current)) ...[
          const SizedBox(height: 10),
          Text('الفولدر الحالي: ${widget.current}', style: const TextStyle(color: C.muted, fontSize: 12), textDirection: TextDirection.ltr),
        ],
      ]),
    );
  }
}

/// ملفات القوايم اللي جاية من تليفون تاني (واتساب / Download / فولدر التطبيق)
Future<List<FoundFile>> listListsFiles(MoneyData d) async {
  final out = <FoundFile>[];
  final seen = <String>{};
  final dirs = <String>[...await scanFolders()];
  try {
    final dir = await getAppFolder(requestPermissions: false);
    dirs.insert(0, '${dir.path}/Lists');
  } catch (_) {}
  for (final p in dirs) {
    try {
      final dd = Directory(p);
      if (!await dd.exists()) continue;
      await for (final e in dd.list()) {
        if (e is! File || !seen.add(e.path)) continue;
        final n = e.uri.pathSegments.last;
        if (!n.endsWith('.json') || !n.startsWith(kListsFilePrefix)) continue;
        final j = await _readJson(e);
        if (j == null || j['app'] != kAppName || j['type'] != 'lists') continue;
        if ('${j['phoneId']}' == d.phoneId) continue;
        final when = DateTime.tryParse('${j['exportedAt'] ?? ''}') ?? await e.lastModified();
        final np = (j['people'] as List?)?.length ?? 0;
        var nc = 0;
        if (j['categories'] is Map) {
          for (final v in (j['categories'] as Map).values) {
            if (v is List) nc += v.length;
          }
        }
        final nb = (j['bills'] as List?)?.length ?? 0;
        out.add(FoundFile(e, when,
            '${j['phoneName'] ?? 'تليفون تاني'} • ${arDateShort(isoDate(when))} ${when.hour.toString().padLeft(2, '0')}:${when.minute.toString().padLeft(2, '0')} • $np فرد • $nc بند • $nb فاتورة',
            j));
      }
    } catch (_) {}
  }
  out.sort((a, b) => b.when.compareTo(a.when));
  return out;
}

/// ملف القوايم ← واتساب (أو أي برنامج)
Future<void> sendListsFile(BuildContext context, MoneyData d) async {
  try {
    Directory out;
    try {
      final dir = await getAppFolder();
      out = Directory('${dir.path}/Lists');
      if (!await out.exists()) await out.create(recursive: true);
    } catch (_) {
      out = await getTemporaryDirectory();
    }
    final now = DateTime.now();
    final stamp = '${isoDate(now)}_${now.hour.toString().padLeft(2, '0')}${now.minute.toString().padLeft(2, '0')}';
    final f = File('${out.path}/$kListsFilePrefix$stamp.json');
    await f.writeAsString(jsonEncode(d.exportLists()), flush: true);
    await Share.shareXFiles([XFile(f.path, mimeType: 'application/json')],
        text: 'A2R Money — القوايم (الأسرة والبنود والفواتير) من غير أي فلوس');
  } catch (e) {
    if (context.mounted) toast(context, 'الملف ماتبعتش: $e', bad: true);
  }
}

/// استلام القوايم: اختار ملف ← معاينة («هيتضاف: …») ← دمج من غير تكرار.
/// بيرجّع true لو اتضاف حاجة (لازم تحفظ بعدها).
Future<bool> receiveListsFlow(BuildContext context, MoneyData d) async {
  await getAppFolder(requestPermissions: true);
  final files = await listListsFiles(d);
  if (!context.mounted) return false;
  final pick = await showModalBottomSheet<FoundFile>(
    context: context,
    isScrollControlled: true,
    builder: (c) => SafeArea(
      child: Padding(
        padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
        child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
          const Text('ملفات القوايم', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
          const Text('من واتساب و Download — افتح الملف في الشات مرة الأول عشان يتنزل', style: TextStyle(color: C.muted, fontSize: 12)),
          const SizedBox(height: 10),
          if (files.isEmpty)
            const Padding(
              padding: EdgeInsets.all(16),
              child: Text('مفيش ملفات قوايم. على التليفون التاني: الإعدادات ← «ابعت القوايم لتليفون تاني» وابعته على واتساب.',
                  style: TextStyle(color: C.muted)),
            ),
          Flexible(
            child: ListView(shrinkWrap: true, children: [
              for (final f in files)
                ListTile(
                  leading: const Icon(Icons.playlist_add_check_rounded, color: C.emerald),
                  title: Text(f.label, style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 13.5)),
                  onTap: () => Navigator.pop(c, f),
                ),
            ]),
          ),
        ]),
      ),
    ),
  );
  if (pick == null || !context.mounted) return false;
  final preview = d.mergeListsJson(pick.json, apply: false, addPlaces: true);
  if (!preview.ok) {
    toast(context, preview.error, bad: true);
    return false;
  }
  if (preview.total == 0) {
    toast(context, 'كل اللي في الملف موجود عندك — مفيش حاجة جديدة تتضاف');
    return false;
  }
  var addPlaces = true;
  final go = await showDialog<bool>(
    context: context,
    builder: (c) => StatefulBuilder(
      builder: (c, set) => AlertDialog(
        title: const Text('تضيف القوايم دي؟', style: TextStyle(fontWeight: FontWeight.w800, fontSize: 18)),
        content: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.start, children: [
          Text('من «${preview.fromName}»', style: const TextStyle(color: C.muted, fontSize: 12.5)),
          const SizedBox(height: 8),
          Text('هيتضاف: ${preview.describe(withPlaces: addPlaces)}', style: const TextStyle(fontWeight: FontWeight.w700)),
          if (preview.missingPlaces > 0)
            CheckboxListTile(
              contentPadding: EdgeInsets.zero,
              value: addPlaces,
              onChanged: (v) => set(() => addPlaces = v ?? true),
              title: Text('ضيف الأماكن كمان (${preview.missingPlaces}) برصيد صفر', style: const TextStyle(fontSize: 13.5)),
            ),
          const SizedBox(height: 6),
          const Text('بيتضاف الناقص بس — مفيش حاجة عندك هتتمسح أو تتغير، والأرصدة والعمليات زي ما هي.',
              style: TextStyle(color: C.muted, fontSize: 12)),
        ]),
        actions: [
          TextButton(onPressed: () => Navigator.pop(c, false), child: const Text('لأ')),
          FilledButton(onPressed: () => Navigator.pop(c, true), child: const Text('ضيفهم')),
        ],
      ),
    ),
  );
  if (go != true || !context.mounted) return false;
  final r = d.mergeListsJson(pick.json, addPlaces: addPlaces);
  if (!r.ok) {
    toast(context, r.error, bad: true);
    return false;
  }
  toast(context, r.total == 0 ? 'مفيش حاجة جديدة' : 'اتضاف: ${r.describe()}');
  return r.total > 0;
}


/// ملفات التحويل اللي جاية من الأساسي (واتساب / Download / فولدر التطبيق)
Future<List<FoundFile>> listTransferFiles(MoneyData d) async {
  final out = <FoundFile>[];
  final seen = <String>{};
  final ids = <String>{};
  final dirs = <String>[...await scanFolders()];
  try {
    final dir = await getAppFolder(requestPermissions: false);
    dirs.insert(0, '${dir.path}/Transfers');
  } catch (_) {}
  for (final p in dirs) {
    try {
      final dd = Directory(p);
      if (!await dd.exists()) continue;
      await for (final e in dd.list()) {
        if (e is! File || !seen.add(e.path)) continue;
        final n = e.uri.pathSegments.last;
        if (!n.endsWith('.json') || !n.startsWith(kTransferFilePrefix)) continue;
        final j = await _readJson(e);
        if (j == null || j['app'] != kAppName || j['type'] != 'transfer') continue;
        if ('${j['fromPhoneId']}' == d.phoneId) continue;
        final id = '${j['id']}';
        if (!ids.add(id)) continue; // نفس التحويل في فولدرين
        final when = DateTime.tryParse('${j['createdAt'] ?? ''}') ?? await e.lastModified();
        final amt = (j['amount'] as num?)?.toInt() ?? 0;
        final done = d.hasReceivedTransfer(id);
        out.add(FoundFile(e, when,
            '${egp(amt)} • ${arDateShort('${j['date']}')}${done ? ' • اتسجل قبل كده' : ''}', j));
      }
    } catch (_) {}
  }
  out.sort((a, b) => b.when.compareTo(a.when));
  return out;
}

Future<void> _prune(Directory b, String prefix, int keep) async {
  final files = <File>[];
  await for (final e in b.list()) {
    if (e is File && e.uri.pathSegments.last.startsWith(prefix)) files.add(e);
  }
  files.sort((x, y) => y.path.compareTo(x.path));
  for (final f in files.skip(keep)) {
    try {
      await f.delete();
    } catch (_) {}
  }
}

String currentInstallId = '';

Map<String, dynamic> backupPayload(MoneyData d) => {
      ...d.toJson(),
      'savedAt': DateTime.now().toIso8601String(),
      'appVersion': kAppVersion,
      'installId': currentInstallId,
    };

/// نسخة تلقائية لكل يوم (بتتحدث طول اليوم)، وبنحتفظ بآخر ١٤ يوم
Future<File?> writeAutoBackup(MoneyData d) async {
  try {
    if (!await hasSharedStorageAccess()) return null;
    final dir = await getAppFolder(requestPermissions: false);
    final b = Directory('${dir.path}/$kBackupsFolder');
    if (!await b.exists()) await b.create(recursive: true);
    final f = File('${b.path}/$kAutoPrefix${isoDate(DateTime.now())}.json');
    if (await f.exists()) {
      final old = await _readJson(f);
      if (old != null && '${old['installId'] ?? ''}' != currentInstallId) {
        final n = DateTime.now();
        await f.rename('${b.path}/$kAutoPrefix${isoDate(n)}_old${n.hour}${n.minute}${n.second}.json');
      }
    }
    final tmp = File('${f.path}.tmp');
    await tmp.writeAsString(jsonEncode(backupPayload(d)), flush: true);
    if (await f.exists()) await f.delete();
    await tmp.rename(f.path);
    await _prune(b, kAutoPrefix, kKeepAutoBackups);
    return f;
  } catch (_) {
    return null;
  }
}

Future<File?> writeSafetyBackup(MoneyData d) async {
  try {
    final dir = await getAppFolder(requestPermissions: false);
    final b = Directory('${dir.path}/$kBackupsFolder');
    if (!await b.exists()) await b.create(recursive: true);
    final n = DateTime.now();
    final f = File('${b.path}/$kBeforeRestorePrefix${isoDate(n)}_${two(n.hour)}${two(n.minute)}${two(n.second)}.json');
    await f.writeAsString(jsonEncode(backupPayload(d)), flush: true);
    await _prune(b, kBeforeRestorePrefix, kKeepSafetyBackups);
    return f;
  } catch (_) {
    return null;
  }
}

class Store {
  static Future<MoneyData?> load() async {
    final sp = await SharedPreferences.getInstance();
    currentInstallId = sp.getString('installId') ?? '';
    if (currentInstallId.isEmpty) {
      currentInstallId = newId('in');
      await sp.setString('installId', currentInstallId);
    }
    final s = sp.getString(kPrefsKey);
    if (s == null) return null;
    try {
      return MoneyData.fromJson(Map<String, dynamic>.from(jsonDecode(s) as Map));
    } catch (_) {
      return null;
    }
  }

  static Future<void> save(MoneyData d) async {
    final sp = await SharedPreferences.getInstance();
    await sp.setString(kPrefsKey, jsonEncode(d.toJson()));
  }
}

const String kXlsxMime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

Future<void> openOrShareFile(BuildContext context, String path, String mime) async {
  OpenResult res;
  try {
    res = await OpenFilex.open(path, type: mime);
  } catch (e) {
    res = OpenResult(type: ResultType.error, message: '$e');
  }
  if (res.type == ResultType.done) return;
  if (context.mounted) {
    ScaffoldMessenger.of(context).showSnackBar(const SnackBar(
        content: Text('مفيش برنامج فتح الملف — هنفتحلك المشاركة (نزّل WPS أو Excel لو عايز تفتحه)')));
  }
  try {
    await Share.shareXFiles([XFile(path, mimeType: mime)]);
  } catch (_) {}
}

// ==================== عناصر الشكل ====================

IconData placeIcon(String kind) {
  switch (kind) {
    case kPlaceCash:
      return Icons.payments_rounded;
    case kPlaceBank:
      return Icons.account_balance_rounded;
    case kPlaceWallet:
      return Icons.account_balance_wallet_rounded;
    default:
      return Icons.savings_rounded;
  }
}

Color placeColor(String kind) {
  switch (kind) {
    case kPlaceCash:
      return C.emerald;
    case kPlaceBank:
      return C.blue;
    case kPlaceWallet:
      return C.gold;
    default:
      return C.violet;
  }
}

IconData typeIcon(String t) {
  switch (t) {
    case tExpense:
      return Icons.shopping_bag_rounded;
    case tAdd:
      return Icons.add_card_rounded;
    case tWithdraw:
      return Icons.outbox_rounded;
    case tTransfer:
      return Icons.swap_horiz_rounded;
    case tAdvance:
      return Icons.handshake_rounded;
    case tRepay:
      return Icons.keyboard_return_rounded;
    case tBill:
      return Icons.receipt_long_rounded;
    case tToPhone:
      return Icons.send_to_mobile_rounded;
    case tFromPhone:
      return Icons.install_mobile_rounded;
  }
  return Icons.circle;
}

Color typeColor(String t) {
  switch (t) {
    case tAdd:
    case tRepay:
    case tFromPhone:
      return C.emerald;
    case tTransfer:
    case tToPhone:
      return C.blue;
    case tAdvance:
      return C.violet;
    case tBill:
      return C.gold;
    default:
      return C.red;
  }
}

class SectionTitle extends StatelessWidget {
  final String text;
  final Widget? trailing;
  const SectionTitle(this.text, {super.key, this.trailing});
  @override
  Widget build(BuildContext context) => Padding(
        padding: const EdgeInsets.fromLTRB(4, 18, 4, 8),
        child: Row(children: [
          Container(width: 4, height: 18, decoration: BoxDecoration(color: C.gold, borderRadius: BorderRadius.circular(4))),
          const SizedBox(width: 8),
          Expanded(child: Text(text, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w800, color: C.text))),
          if (trailing != null) trailing!,
        ]),
      );
}

/// البادج اللي فوق: اسم التليفون ده. الضغط عليه = تغيير الاسم.
class PhoneBadge extends StatelessWidget {
  final MoneyData data;
  final VoidCallback onTap;
  const PhoneBadge({super.key, required this.data, required this.onTap});
  @override
  Widget build(BuildContext context) {
    const color = C.gold;
    return Tooltip(
      message: 'اسم التليفون ده — دوس عشان تغيّره',
      child: InkWell(
        borderRadius: BorderRadius.circular(20),
        onTap: onTap,
        child: Container(
          constraints: const BoxConstraints(maxWidth: 150),
          padding: const EdgeInsets.symmetric(horizontal: 9, vertical: 4),
          decoration: BoxDecoration(
              color: color.withValues(alpha: 0.16),
              borderRadius: BorderRadius.circular(20),
              border: Border.all(color: color.withValues(alpha: 0.5))),
          child: Row(mainAxisSize: MainAxisSize.min, children: [
            const Icon(Icons.phone_android_rounded, size: 14, color: color),
            const SizedBox(width: 4),
            Flexible(
              child: Text(data.badgeName,
                  maxLines: 1,
                  overflow: TextOverflow.ellipsis,
                  style: const TextStyle(fontSize: 11.5, color: color, fontWeight: FontWeight.w700)),
            ),
          ]),
        ),
      ),
    );
  }
}

class Pill extends StatelessWidget {
  final String text;
  final Color color;
  const Pill(this.text, {super.key, this.color = C.gold});
  @override
  Widget build(BuildContext context) => Container(
        padding: const EdgeInsets.symmetric(horizontal: 9, vertical: 3),
        decoration: BoxDecoration(
            color: color.withValues(alpha: 0.16),
            borderRadius: BorderRadius.circular(20),
            border: Border.all(color: color.withValues(alpha: 0.5))),
        child: Text(text, style: TextStyle(fontSize: 11.5, color: color, fontWeight: FontWeight.w700)),
      );
}

class MonthSwitcher extends StatelessWidget {
  final String month;
  final ValueChanged<String> onChanged;
  const MonthSwitcher({super.key, required this.month, required this.onChanged});
  @override
  Widget build(BuildContext context) {
    final now = monthKeyOf(DateTime.now());
    final f = monthFirst(month);
    return Container(
      margin: const EdgeInsets.only(bottom: 6),
      decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(16), border: Border.all(color: C.line)),
      child: Row(children: [
        IconButton(tooltip: 'الشهر اللي فات', onPressed: () => onChanged(shiftMonth(month, -1)), icon: const Icon(Icons.chevron_right_rounded)),
        Expanded(
          child: InkWell(
            onTap: month == now ? null : () => onChanged(now),
            child: Column(children: [
              Text(monthLabel(month), style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w800)),
              Text(month == now ? 'الشهر الحالي (من ١ لـ ${daysInMonth(f.year, f.month)})' : 'دوس هنا ترجع للشهر الحالي',
                  style: const TextStyle(fontSize: 11, color: C.muted)),
            ]),
          ),
        ),
        IconButton(tooltip: 'الشهر الجاي', onPressed: () => onChanged(shiftMonth(month, 1)), icon: const Icon(Icons.chevron_left_rounded)),
      ]),
    );
  }
}

Future<bool> confirmDialog(BuildContext context, String title, String body,
    {String yes = 'أيوه', String no = 'لأ', bool danger = false}) async {
  final r = await showDialog<bool>(
    context: context,
    builder: (c) => AlertDialog(
      title: Text(title, style: const TextStyle(fontWeight: FontWeight.w800, fontSize: 18)),
      content: body.isEmpty ? null : Text(body),
      actions: [
        TextButton(onPressed: () => Navigator.pop(c, false), child: Text(no)),
        FilledButton(
            style: danger ? FilledButton.styleFrom(backgroundColor: C.red, foregroundColor: Colors.white) : null,
            onPressed: () => Navigator.pop(c, true),
            child: Text(yes)),
      ],
    ),
  );
  return r == true;
}

void toast(BuildContext context, String m, {bool bad = false}) {
  ScaffoldMessenger.of(context).hideCurrentSnackBar();
  ScaffoldMessenger.of(context).showSnackBar(SnackBar(
    content: Text(m),
    backgroundColor: bad ? const Color(0xFF7F1D1D) : C.card2,
    duration: Duration(seconds: bad ? 5 : 3),
  ));
}

Future<String?> askText(BuildContext context, String title, {String initial = '', String hint = '', bool allowEmpty = false}) async {
  final ctl = TextEditingController(text: initial);
  final r = await showDialog<String>(
    context: context,
    builder: (c) => AlertDialog(
      title: Text(title, style: const TextStyle(fontWeight: FontWeight.w800, fontSize: 18)),
      content: TextField(
        controller: ctl,
        autofocus: true,
        decoration: InputDecoration(hintText: hint),
        onSubmitted: (v) => Navigator.pop(c, v),
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(c), child: const Text('إلغاء')),
        FilledButton(onPressed: () => Navigator.pop(c, ctl.text), child: const Text('حفظ')),
      ],
    ),
  );
  if (r == null) return null;
  if (!allowEmpty && r.trim().isEmpty) return null;
  return r.trim();
}

// ---------- التقويم العربي (السبت أول الأسبوع) ----------

const List<String> kCalHeads = ['سبت', 'حد', 'اتنين', 'تلات', 'أربع', 'خميس', 'جمعة'];

/// مكان اليوم الأول في الشهر لو السبت أول عمود (السبت=0 … الجمعة=6)
int saturdayFirstLead(DateTime firstOfMonth) => (firstOfMonth.weekday + 1) % 7;

Future<DateTime?> pickArabicDate(BuildContext context, DateTime initial, {bool Function(DateTime d)? hasRecords}) =>
    showDialog<DateTime>(context: context, builder: (_) => MoneyCalendar(initial: initial, hasRecords: hasRecords));

class MoneyCalendar extends StatefulWidget {
  final DateTime initial;
  final bool Function(DateTime d)? hasRecords;
  const MoneyCalendar({super.key, required this.initial, this.hasRecords});
  @override
  State<MoneyCalendar> createState() => _MoneyCalendarState();
}

class _MoneyCalendarState extends State<MoneyCalendar> {
  late DateTime picked;
  late DateTime month;
  @override
  void initState() {
    super.initState();
    picked = DateTime(widget.initial.year, widget.initial.month, widget.initial.day);
    month = DateTime(picked.year, picked.month);
  }

  @override
  Widget build(BuildContext context) {
    final now = DateTime.now();
    final today = DateTime(now.year, now.month, now.day);
    final dim = daysInMonth(month.year, month.month);
    final lead = saturdayFirstLead(DateTime(month.year, month.month, 1));
    final cells = <DateTime?>[
      for (var i = 0; i < lead; i++) null,
      for (var d = 1; d <= dim; d++) DateTime(month.year, month.month, d),
    ];
    while (cells.length % 7 != 0) {
      cells.add(null);
    }
    return Dialog(
      insetPadding: const EdgeInsets.symmetric(horizontal: 14, vertical: 24),
      child: SingleChildScrollView(
        padding: const EdgeInsets.fromLTRB(14, 16, 14, 10),
        child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
          const Text('اختار اليوم', style: TextStyle(fontSize: 13, color: C.muted)),
          Text('${arabicWeekdays[picked.weekday - 1]} ${arDate(picked)}',
              style: const TextStyle(fontSize: 21, fontWeight: FontWeight.w800, color: C.gold)),
          const Divider(height: 20, color: C.line),
          Row(children: [
            IconButton(onPressed: () => setState(() => month = DateTime(month.year, month.month - 1)), icon: const Icon(Icons.chevron_right_rounded)),
            Expanded(
                child: Text('${arabicMonths[month.month - 1]} ${month.year}',
                    textAlign: TextAlign.center, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w800))),
            IconButton(onPressed: () => setState(() => month = DateTime(month.year, month.month + 1)), icon: const Icon(Icons.chevron_left_rounded)),
          ]),
          Row(
              children: kCalHeads
                  .map((h) => Expanded(
                      child: Center(
                          child: Text(h,
                              style: TextStyle(
                                  fontSize: 12, fontWeight: FontWeight.w700, color: h == 'جمعة' ? C.muted.withValues(alpha: 0.6) : C.muted)))))
                  .toList()),
          const SizedBox(height: 6),
          for (var w = 0; w < cells.length ~/ 7; w++)
            Row(children: [for (var i = 0; i < 7; i++) Expanded(child: _cell(cells[w * 7 + i], today))]),
          const SizedBox(height: 6),
          Row(children: [
            TextButton(
                onPressed: () => setState(() {
                      picked = today;
                      month = DateTime(today.year, today.month);
                    }),
                child: const Text('النهارده')),
            const Spacer(),
            TextButton(onPressed: () => Navigator.pop(context), child: const Text('إلغاء')),
            const SizedBox(width: 4),
            FilledButton(onPressed: () => Navigator.pop(context, picked), child: const Text('موافق')),
          ]),
        ]),
      ),
    );
  }

  Widget _cell(DateTime? d, DateTime today) {
    if (d == null) return const SizedBox(height: 44);
    final sel = d == picked;
    final isToday = d == today;
    final has = widget.hasRecords?.call(d) ?? false;
    return SizedBox(
      height: 44,
      child: Center(
        child: InkWell(
          customBorder: const CircleBorder(),
          onTap: () => setState(() => picked = d),
          child: Container(
            width: 38,
            height: 38,
            alignment: Alignment.center,
            decoration: BoxDecoration(
              shape: BoxShape.circle,
              color: sel ? C.gold : (has ? C.emerald.withValues(alpha: 0.18) : null),
              border: Border.all(
                  color: sel ? C.gold : (isToday ? C.text : (has ? C.emerald.withValues(alpha: 0.6) : Colors.transparent)),
                  width: isToday ? 2 : 1.2),
            ),
            child: Text('${d.day}',
                style: TextStyle(
                    fontSize: 15,
                    fontWeight: sel || isToday ? FontWeight.w800 : FontWeight.normal,
                    color: sel ? const Color(0xFF231A00) : C.text)),
          ),
        ),
      ),
    );
  }
}

// ---------- اختيار المكان (إجباري، وظاهر رصيده) ----------

class PlaceField extends StatelessWidget {
  final MoneyData data;
  final String? placeId;
  final String label;
  final bool incoming;
  final String? errorText;
  final String? excludeId;
  final String? replacingTxnId;
  final ValueChanged<String> onPicked;
  const PlaceField({
    super.key,
    required this.data,
    required this.placeId,
    required this.label,
    required this.onPicked,
    this.incoming = false,
    this.errorText,
    this.excludeId,
    this.replacingTxnId,
  });

  @override
  Widget build(BuildContext context) {
    final p = placeId == null ? null : data.placeById(placeId!);
    final bal = p == null ? 0 : data.balanceOf(p.id, excludeTxnId: replacingTxnId);
    final err = errorText != null;
    return Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
      Row(children: [
        Icon(incoming ? Icons.south_west_rounded : Icons.north_east_rounded, size: 16, color: incoming ? C.emerald : C.red),
        const SizedBox(width: 4),
        Text(label, style: TextStyle(fontSize: 13.5, fontWeight: FontWeight.w800, color: err ? C.red : C.text)),
      ]),
      const SizedBox(height: 6),
      InkWell(
        borderRadius: BorderRadius.circular(16),
        onTap: () async {
          final id = await pickPlace(context, data, title: label, excludeId: excludeId, replacingTxnId: replacingTxnId);
          if (id != null) onPicked(id);
        },
        child: Container(
          padding: const EdgeInsets.symmetric(horizontal: 14, vertical: 12),
          decoration: BoxDecoration(
            color: C.bg2,
            borderRadius: BorderRadius.circular(16),
            border: Border.all(color: err ? C.red : (p == null ? C.line : placeColor(p.kind)), width: 1.4),
          ),
          child: Row(children: [
            Icon(p == null ? Icons.help_outline_rounded : placeIcon(p.kind), color: p == null ? C.muted : placeColor(p.kind)),
            const SizedBox(width: 10),
            Expanded(
              child: Text(p?.name ?? 'دوس واختار المكان',
                  style: TextStyle(fontSize: 16, fontWeight: FontWeight.w800, color: p == null ? C.muted : C.text)),
            ),
            if (p != null)
              Column(crossAxisAlignment: CrossAxisAlignment.end, children: [
                const Text('الرصيد دلوقتي', style: TextStyle(fontSize: 10.5, color: C.muted)),
                Text(egp(bal), style: TextStyle(fontWeight: FontWeight.w800, color: bal < 0 ? C.red : C.emerald)),
              ]),
            const SizedBox(width: 6),
            const Icon(Icons.expand_more_rounded, color: C.muted),
          ]),
        ),
      ),
      if (err)
        Padding(padding: const EdgeInsets.only(top: 4, right: 6), child: Text(errorText!, style: const TextStyle(color: C.red, fontSize: 12))),
    ]);
  }
}

Future<String?> pickPlace(BuildContext context, MoneyData d, {required String title, String? excludeId, String? replacingTxnId}) {
  final list = d.activePlaces.where((p) => p.id != excludeId).toList();
  return showModalBottomSheet<String>(
    context: context,
    isScrollControlled: true,
    builder: (c) => SafeArea(
      child: Padding(
        padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
        child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
          Text(title, style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
          const SizedBox(height: 10),
          if (list.isEmpty)
            const Padding(padding: EdgeInsets.all(12), child: Text('مفيش أماكن — ضيف مكان من الإعدادات', style: TextStyle(color: C.muted))),
          Flexible(
            child: ListView(shrinkWrap: true, children: [
              for (final p in list)
                Padding(
                  padding: const EdgeInsets.only(bottom: 8),
                  child: ListTile(
                    shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16), side: const BorderSide(color: C.line)),
                    tileColor: C.card,
                    leading: CircleAvatar(
                        backgroundColor: placeColor(p.kind).withValues(alpha: 0.18), child: Icon(placeIcon(p.kind), color: placeColor(p.kind))),
                    title: Text(p.name, style: const TextStyle(fontWeight: FontWeight.w800)),
                    trailing: Builder(builder: (_) {
                      final b = d.balanceOf(p.id, excludeTxnId: replacingTxnId);
                      return Text(egp(b), style: TextStyle(fontWeight: FontWeight.w800, color: b < 0 ? C.red : C.emerald));
                    }),
                    onTap: () => Navigator.pop(c, p.id),
                  ),
                ),
            ]),
          ),
        ]),
      ),
    ),
  );
}

// ---------- كارت عملية ----------

class TxnTile extends StatelessWidget {
  final MoneyData data;
  final Txn t;
  final VoidCallback? onTap;
  final VoidCallback? onLongPress;
  const TxnTile({super.key, required this.data, required this.t, this.onTap, this.onLongPress});

  String title() {
    switch (t.type) {
      case tExpense:
        return t.category.isEmpty ? 'مصروف ${t.group}' : t.category;
      case tBill:
        return 'فاتورة ${t.category}';
      case tWithdraw:
        return t.category.isNotEmpty ? 'سحب • ${t.category}' : 'سحب';
      case tAdd:
        return 'إضافة رصيد';
      case tTransfer:
        return 'تحويل';
      case tAdvance:
        return 'سلفة لـ ${data.personName(t.personId)}';
      case tRepay:
        return 'رد سلفة من ${data.personName(t.personId)}';
      case tToPhone:
        return 'اتحول لـ ${data.otherName}';
      case tFromPhone:
        return 'استلام من ${data.otherName}';
    }
    return t.type;
  }

  @override
  Widget build(BuildContext context) {
    final col = typeColor(t.type);
    final sub = <String>[
      arDateShort(t.date),
      if (t.type == tTransfer)
        '${data.placeName(t.placeId)} ← ${data.placeName(t.toPlaceId)}'
      else
        (isOut(t.type) ? 'من ${data.placeName(t.placeId)}' : 'لـ ${data.placeName(t.placeId)}'),
      if ((t.type == tExpense || t.type == tWithdraw) && t.group.isNotEmpty) t.group,
      if ((t.type == tExpense || t.type == tWithdraw || t.type == tBill) && t.personId.isNotEmpty) data.personName(t.personId),
    ];
    final inc = t.type == tAdd || t.type == tRepay || t.type == tFromPhone;
    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Material(
        color: C.card,
        borderRadius: BorderRadius.circular(18),
        child: InkWell(
          borderRadius: BorderRadius.circular(18),
          onTap: onTap,
          onLongPress: onLongPress,
          child: Container(
            padding: const EdgeInsets.all(12),
            decoration: BoxDecoration(borderRadius: BorderRadius.circular(18), border: Border.all(color: C.line)),
            child: Row(children: [
              Container(
                width: 42,
                height: 42,
                decoration: BoxDecoration(color: col.withValues(alpha: 0.16), borderRadius: BorderRadius.circular(13)),
                child: Icon(typeIcon(t.type), color: col),
              ),
              const SizedBox(width: 12),
              Expanded(
                child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                  Text(title(), style: const TextStyle(fontSize: 15, fontWeight: FontWeight.w800), maxLines: 1, overflow: TextOverflow.ellipsis),
                  const SizedBox(height: 2),
                  Text(sub.join(' • '), style: const TextStyle(fontSize: 12, color: C.muted), maxLines: 1, overflow: TextOverflow.ellipsis),
                  if (t.note.isNotEmpty) Text(t.note, style: const TextStyle(fontSize: 12, color: C.text), maxLines: 1, overflow: TextOverflow.ellipsis),
                ]),
              ),
              const SizedBox(width: 8),
              Text(
                t.type == tTransfer ? money(t.amount) : (inc ? '+${money(t.amount)}' : '-${money(t.amount)}'),
                style: TextStyle(fontSize: 16, fontWeight: FontWeight.w900, color: t.type == tTransfer ? C.blue : (inc ? C.emerald : C.text)),
              ),
            ]),
          ),
        ),
      ),
    );
  }
}

// ==================== تسجيل / تعديل عملية ====================

class TxnFormScreen extends StatefulWidget {
  final MoneyData data;
  final String type;
  final Txn? existing;
  final Bill? bill; // دفع فاتورة
  final String? billMonth;
  final String? presetPersonId;
  const TxnFormScreen({
    super.key,
    required this.data,
    required this.type,
    this.existing,
    this.bill,
    this.billMonth,
    this.presetPersonId,
  });
  @override
  State<TxnFormScreen> createState() => _TxnFormScreenState();
}

/// نتيجة الشاشة: عملية اتحفظت، أو «delete» لو اتمسحت
class TxnFormResult {
  final Txn? saved;
  final bool deleted;
  final bool resend; // ابعت ملف التحويل تاني
  const TxnFormResult({this.saved, this.deleted = false, this.resend = false});
}

class _TxnFormScreenState extends State<TxnFormScreen> {
  final amountCtl = TextEditingController();
  final noteCtl = TextEditingController();
  late String type;
  String? placeId;
  String? toPlaceId;
  String group = '';
  String category = '';
  String personId = '';
  late DateTime date;
  bool withdrawOnItem = false;
  bool typingCat = false; // «✏️ اكتب بند جديد»
  final catCtl = TextEditingController();
  String? amountErr, placeErr, toErr, groupErr, catErr, personErr;

  MoneyData get d => widget.data;
  bool get editing => widget.existing != null;
  bool get outgoing => isOut(type);

  @override
  void initState() {
    super.initState();
    type = widget.type;
    final e = widget.existing;
    if (e != null) {
      amountCtl.text = money(e.amount).replaceAll(',', '');
      noteCtl.text = e.note;
      placeId = d.placeById(e.placeId) != null ? e.placeId : null;
      toPlaceId = e.toPlaceId.isEmpty ? null : e.toPlaceId;
      group = e.group;
      category = e.category;
      personId = e.personId;
      date = parseIso(e.date);
      withdrawOnItem = type == tWithdraw && e.category.isNotEmpty;
    } else {
      date = DateTime.now();
      // آخر مكان اتستخدم بيتختار لوحده — بس بيبان واضح وتقدر تغيره
      final last = outgoing ? d.lastPlaceId : d.lastInPlaceId;
      final lp = d.placeById(last);
      placeId = (lp != null && !lp.archived) ? lp.id : null;
      if (type == tExpense) group = kGroupPersonal;
      personId = widget.presetPersonId ?? '';
      final b = widget.bill;
      if (b != null) {
        group = kGroupHome;
        category = b.name;
        if (b.usual > 0) amountCtl.text = money(b.usual).replaceAll(',', '');
      }
    }
  }

  String get title {
    if (widget.bill != null) return 'دفع فاتورة ${widget.bill!.name}';
    final l = type == tToPhone ? 'تحويل لـ ${d.otherName}' : (kTypeLabels[type] ?? type);
    return editing ? 'تعديل $l' : l;
  }

  String get placeLabel {
    switch (type) {
      case tAdd:
        return 'الفلوس اتضافت لـ فين؟';
      case tRepay:
        return 'الفلوس اللي رجعت دخلت فين؟';
      case tFromPhone:
        return 'الفلوس دخلت فين؟';
      case tToPhone:
        return 'هتتحول من أنهي مكان؟';
      case tTransfer:
        return 'من مكان';
      default:
        return 'اتسحب منين؟';
    }
  }

  bool get usesCategory => type == tExpense || (type == tWithdraw && withdrawOnItem);

  Future<void> save() async {
    final amt = parseMoney(amountCtl.text);
    if (usesCategory && typingCat) category = cleanLabel(catCtl.text);
    setState(() {
      amountErr = (amt == null || amt <= 0) ? 'اكتب المبلغ' : null;
      placeErr = placeId == null ? 'لازم تختار المكان' : null;
      toErr = type == tTransfer
          ? (toPlaceId == null
              ? 'اختار المكان اللي الفلوس رايحة له'
              : (toPlaceId == placeId ? 'لازم مكانين مختلفين' : null))
          : null;
      groupErr = type == tExpense && group.isEmpty ? 'اختار النوع' : null;
      catErr = (type == tExpense && category.isEmpty) || (type == tWithdraw && withdrawOnItem && category.isEmpty)
          ? (typingCat ? 'اكتب اسم البند الجديد أو اختار من اللي فوق' : 'اختار البند')
          : null;
      personErr = (type == tAdvance || type == tRepay) && personId.isEmpty ? 'اختار الشخص' : null;
    });
    if ([amountErr, placeErr, toErr, groupErr, catErr, personErr].any((e) => e != null)) return;
    final t = Txn(
      id: widget.existing?.id ?? newId('tx'),
      type: type,
      amount: amt!,
      date: isoDate(date),
      placeId: placeId!,
      toPlaceId: type == tTransfer ? toPlaceId! : '',
      group: type == tExpense || type == tBill || (type == tWithdraw && withdrawOnItem) ? group : '',
      category: type == tExpense || type == tBill || (type == tWithdraw && withdrawOnItem) ? category : '',
      personId: (type == tExpense || type == tAdvance || type == tRepay || type == tWithdraw || type == tBill)
          ? personId
          : '',
      billId: widget.bill?.id ?? widget.existing?.billId ?? '',
      billMonth: widget.billMonth ?? widget.existing?.billMonth ?? '',
      note: noteCtl.text.trim(),
      refId: widget.existing?.refId ?? '',
      createdAt: widget.existing?.createdAt,
    );
    // تحذير بس (مش منع) لو مكان هيبقى بالسالب
    final warn = d.negativeWarnings(t, replacingId: widget.existing?.id);
    if (warn.isNotEmpty) {
      final go = await confirmDialog(context, 'الرصيد هيبقى بالسالب',
          '${warn.join('\n')}\n\nتسجّل برضه؟',
          yes: 'سجّل برضه', no: 'ارجع');
      if (!go) return;
    }
    if (type == tRepay) {
      final rem = d.advanceRemaining(personId) + (widget.existing?.type == tRepay ? widget.existing!.amount : 0);
      if (amt > rem && mounted) {
        final go = await confirmDialog(context, 'المبلغ أكبر من الباقي',
            'الباقي على ${d.personName(personId)}: ${egp(rem)}\nالرد: ${egp(amt)}\n\nتسجّل برضه؟',
            yes: 'سجّل برضه', no: 'ارجع');
        if (!go) return;
      }
    }
    if (!mounted) return;
    // اللي اتكتب بإيد بيتحفظ لوحده ويبقى اختيار المرة الجاية (من غير تكرار)
    if (t.category.isNotEmpty && (t.type == tExpense || t.type == tWithdraw)) {
      t.category = d.addCategory(t.group, t.category);
    }
    if (t.note.isNotEmpty && t.type == tAdd) d.addSource(t.note);
    if (t.note.isNotEmpty && t.type == tWithdraw) d.addReason(t.note);
    Navigator.pop(context, TxnFormResult(saved: t));
  }

  Widget chips(List<String> items, String sel, ValueChanged<String> on, {Color color = C.gold}) => Wrap(
        spacing: 8,
        runSpacing: 8,
        children: [
          for (final x in items)
            ChoiceChip(
              label: Text(x),
              selected: sel == x,
              selectedColor: color.withValues(alpha: 0.28),
              side: BorderSide(color: sel == x ? color : C.line),
              labelStyle: TextStyle(fontWeight: FontWeight.w700, color: sel == x ? C.text : C.muted),
              showCheckmark: false,
              onSelected: (_) => on(x),
            ),
        ],
      );

  void startTypingCat() => setState(() {
        typingCat = true;
        category = '';
        catErr = null;
      });

  Widget categoryPicker(List<String> cats) {
    final typed = cleanLabel(catCtl.text);
    String? existing;
    if (typed.isNotEmpty) {
      for (final c in cats) {
        if (sameLabel(c, typed)) existing = c;
      }
    }
    return Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
      Wrap(spacing: 8, runSpacing: 8, children: [
        for (final x in cats)
          ChoiceChip(
            label: Text(x),
            selected: !typingCat && category == x,
            selectedColor: C.emerald.withValues(alpha: 0.28),
            side: BorderSide(color: !typingCat && category == x ? C.emerald : C.line),
            labelStyle: TextStyle(fontWeight: FontWeight.w700, color: !typingCat && category == x ? C.text : C.muted),
            showCheckmark: false,
            onSelected: (_) => setState(() {
              typingCat = false;
              category = x;
              catErr = null;
            }),
          ),
        ChoiceChip(
          avatar: const Icon(Icons.edit_rounded, size: 16, color: C.gold),
          label: const Text('اكتب بند جديد'),
          selected: typingCat,
          selectedColor: C.gold.withValues(alpha: 0.25),
          side: BorderSide(color: typingCat ? C.gold : C.gold.withValues(alpha: 0.5)),
          labelStyle: const TextStyle(fontWeight: FontWeight.w800, color: C.gold),
          showCheckmark: false,
          onSelected: (_) => startTypingCat(),
        ),
      ]),
      if (typingCat) ...[
        const SizedBox(height: 10),
        TextField(
          controller: catCtl,
          autofocus: true,
          textInputAction: TextInputAction.done,
          onChanged: (_) => setState(() => catErr = null),
          decoration: InputDecoration(
            labelText: 'اسم البند الجديد',
            hintText: 'مثلاً: جيم',
            helperText: existing != null
                ? 'موجود قبل كده باسم «$existing» — هيتسجل عليه (مش هيتكرر)'
                : 'هيتحفظ لوحده في بنود «$group» ويطلعلك المرة الجاية',
            helperMaxLines: 2,
            prefixIcon: const Icon(Icons.edit_note_rounded, color: C.gold),
          ),
        ),
      ],
    ]);
  }

  /// اقتراحات (مصدر الفلوس / سبب السحب) اللي اتكتبت قبل كده
  Widget noteSuggestions(List<String> items) {
    if (items.isEmpty) return const SizedBox.shrink();
    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Wrap(spacing: 6, runSpacing: 6, children: [
        for (final x in items.reversed.take(12))
          ActionChip(
            label: Text(x),
            onPressed: () => setState(() => noteCtl.text = x),
          ),
      ]),
    );
  }

  Future<void> pickPerson({bool familyOnly = false}) async {
    final list = d.people.where((p) => !familyOnly || p.family).toList();
    final r = await showModalBottomSheet<String>(
      context: context,
      isScrollControlled: true,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            Text(familyOnly ? 'المصروف ده لمين؟' : 'الشخص', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            const SizedBox(height: 10),
            Flexible(
              child: ListView(shrinkWrap: true, children: [
                if (familyOnly)
                  ListTile(
                    leading: const Icon(Icons.block_rounded, color: C.muted),
                    title: const Text('مش لحد معين'),
                    onTap: () => Navigator.pop(c, ''),
                  ),
                for (final p in list)
                  ListTile(
                    leading: Icon(p.child ? Icons.child_care_rounded : Icons.person_rounded,
                        color: p.child ? C.gold : C.emerald),
                    title: Text(p.name, style: const TextStyle(fontWeight: FontWeight.w700)),
                    subtitle: (type == tAdvance || type == tRepay)
                        ? Text('الباقي عليه: ${egp(d.advanceRemaining(p.id))}', style: const TextStyle(color: C.muted))
                        : null,
                    onTap: () => Navigator.pop(c, p.id),
                  ),
                ListTile(
                  leading: const Icon(Icons.person_add_alt_1_rounded, color: C.gold),
                  title: const Text('اكتب اسم جديد', style: TextStyle(color: C.gold, fontWeight: FontWeight.w700)),
                  onTap: () => Navigator.pop(c, '__new__'),
                ),
              ]),
            ),
          ]),
        ),
      ),
    );
    if (r == null) return;
    if (r == '__new__') {
      if (!mounted) return;
      final n = await askText(context, familyOnly ? 'اسم فرد الأسرة' : 'اسم الشخص',
          hint: 'هيتحفظ ويطلعلك في الاختيارات المرة الجاية');
      if (n == null) return;
      // لو الاسم موجود قبل كده بيتختار هو (مش بيتكرر)
      final p = d.findOrAddPerson(n, family: familyOnly);
      if (p == null) return;
      if (familyOnly) p.family = true;
      setState(() => personId = p.id);
      return;
    }
    setState(() => personId = r);
  }

  @override
  Widget build(BuildContext context) {
    final col = typeColor(type);
    final cats = d.categories[group] ?? const <String>[];
    return Scaffold(
      appBar: AppBar(
        title: Text(title),
        actions: [
          if (editing)
            IconButton(
              tooltip: 'مسح',
              icon: const Icon(Icons.delete_outline_rounded, color: C.red),
              onPressed: () async {
                final ok = await confirmDialog(context, 'مسح العملية؟',
                    '${kTypeLabels[type]} ${egp(widget.existing!.amount)}\nالرصيد هيرجع زي ما كان.',
                    yes: 'امسح', danger: true);
                if (ok && context.mounted) Navigator.pop(context, const TxnFormResult(deleted: true));
              },
            ),
        ],
      ),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 8, 16, 120), children: [
        Container(
          padding: const EdgeInsets.all(16),
          decoration: BoxDecoration(
            borderRadius: BorderRadius.circular(22),
            gradient: LinearGradient(colors: [col.withValues(alpha: 0.28), C.card]),
            border: Border.all(color: col.withValues(alpha: 0.5)),
          ),
          child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            Text('المبلغ', style: TextStyle(color: col, fontWeight: FontWeight.w700)),
            TextField(
              controller: amountCtl,
              autofocus: !editing,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              inputFormatters: [FilteringTextInputFormatter.allow(RegExp(r'[0-9٠-٩۰-۹.,٫،]'))],
              style: const TextStyle(fontSize: 34, fontWeight: FontWeight.w900),
              decoration: InputDecoration(
                hintText: '0',
                suffixText: 'جنيه',
                errorText: amountErr,
                filled: false,
                border: InputBorder.none,
                enabledBorder: InputBorder.none,
                focusedBorder: InputBorder.none,
              ),
            ),
          ]),
        ),
        const SizedBox(height: 16),
        PlaceField(
          data: d,
          placeId: placeId,
          label: placeLabel,
          incoming: !outgoing,
          errorText: placeErr,
          replacingTxnId: widget.existing?.id,
          onPicked: (id) => setState(() {
            placeId = id;
            placeErr = null;
          }),
        ),
        if (type == tTransfer) ...[
          const SizedBox(height: 12),
          PlaceField(
            data: d,
            placeId: toPlaceId,
            label: 'لـ مكان',
            incoming: true,
            errorText: toErr,
            excludeId: placeId,
            replacingTxnId: widget.existing?.id,
            onPicked: (id) => setState(() {
              toPlaceId = id;
              toErr = null;
            }),
          ),
        ],
        if (type == tExpense) ...[
          const SectionTitle('النوع'),
          chips(kGroups, group, (g) => setState(() {
                group = g;
                category = '';
                groupErr = null;
              })),
          if (groupErr != null) Text(groupErr!, style: const TextStyle(color: C.red, fontSize: 12)),
        ],
        if (type == tWithdraw) ...[
          const SizedBox(height: 12),
          SwitchListTile(
            contentPadding: EdgeInsets.zero,
            value: withdrawOnItem,
            title: const Text('أحسبه على بند؟', style: TextStyle(fontWeight: FontWeight.w700)),
            subtitle: const Text('اختياري — لو لأ بيتحسب «سحب» لوحده في ملخص الشهر',
                style: TextStyle(color: C.muted, fontSize: 12)),
            onChanged: (v) => setState(() {
              withdrawOnItem = v;
              if (v && group.isEmpty) group = kGroupPersonal;
            }),
          ),
          if (withdrawOnItem) chips(kGroups, group, (g) => setState(() {
                group = g;
                category = '';
              })),
        ],
        if (type == tExpense || (type == tWithdraw && withdrawOnItem)) ...[
          const SectionTitle('البند'),
          categoryPicker(cats),
          if (catErr != null) Padding(padding: const EdgeInsets.only(top: 4), child: Text(catErr!, style: const TextStyle(color: C.red, fontSize: 12))),
        ],
        if (type == tExpense || type == tBill || type == tWithdraw) ...[
          const SectionTitle('لمين؟ (اختياري)'),
          OutlinedButton.icon(
            onPressed: () => pickPerson(familyOnly: true),
            icon: const Icon(Icons.family_restroom_rounded),
            label: Text(personId.isEmpty ? 'مش لحد معين' : d.personName(personId)),
          ),
        ],
        if (type == tAdvance || type == tRepay) ...[
          SectionTitle(type == tAdvance ? 'السلفة لمين؟' : 'مين اللي رجّع؟'),
          OutlinedButton.icon(
            style: OutlinedButton.styleFrom(side: BorderSide(color: personErr != null ? C.red : C.line)),
            onPressed: () => pickPerson(),
            icon: const Icon(Icons.person_rounded),
            label: Text(personId.isEmpty ? 'اختار الشخص' : '${d.personName(personId)} • الباقي ${egp(d.advanceRemaining(personId))}'),
          ),
          if (personErr != null) Text(personErr!, style: const TextStyle(color: C.red, fontSize: 12)),
        ],
        const SectionTitle('اليوم'),
        OutlinedButton.icon(
          onPressed: () async {
            final x = await pickArabicDate(context, date);
            if (x != null) setState(() => date = x);
          },
          icon: const Icon(Icons.calendar_month_rounded),
          label: Text('${arabicWeekdays[date.weekday - 1]} ${arDate(date)}'),
        ),
        const SizedBox(height: 14),
        // «من التليفون التاني» دايمًا موجود كاختيار (لو التحويل اتسجل بإيد من غير ملف)
        if (type == tAdd)
          noteSuggestions([
            ...d.sources.where((x) => !isPhoneSource(x)),
            kFromMainSource,
          ]),
        if (type == tWithdraw) noteSuggestions(d.reasons),
        if (type == tToPhone) ...[
          Container(
            padding: const EdgeInsets.all(12),
            margin: const EdgeInsets.only(bottom: 12),
            decoration: BoxDecoration(
                color: C.blue.withValues(alpha: 0.10),
                borderRadius: BorderRadius.circular(14),
                border: Border.all(color: C.blue.withValues(alpha: 0.4))),
            child: Text(
                'الفلوس دي بتتخصم من المكان هنا، ومش بتتحسب صرف. بعد ما تسجّل هتبعت ملف صغير على واتساب لـ ${d.otherName}، '
                'وهو يدوس «تحديث» (أو «استلم تحويل») فتدخل عنده.',
                style: const TextStyle(fontSize: 12.5)),
          ),
        ],
        TextField(
          controller: noteCtl,
          maxLines: 2,
          decoration: InputDecoration(
            labelText: type == tAdd
                ? 'جاية منين؟ (اختياري)'
                : type == tWithdraw
                    ? 'السبب (اختياري)'
                    : 'ملاحظة (اختياري)',
          ),
        ),
        if (editing && type == tToPhone) ...[
          const SizedBox(height: 14),
          OutlinedButton.icon(
            onPressed: () => Navigator.pop(context, const TxnFormResult(resend: true)),
            icon: const Icon(Icons.send_rounded),
            label: const Text('ابعت ملف التحويل تاني'),
          ),
        ],
      ]),
      bottomNavigationBar: SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 6, 16, 12),
          child: FilledButton.icon(
            style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 16)),
            onPressed: save,
            icon: const Icon(Icons.check_rounded),
            label: Text(editing ? 'حفظ التعديل' : 'سجّل', style: const TextStyle(fontSize: 17, fontWeight: FontWeight.w800)),
          ),
        ),
      ),
    );
  }
}

// ==================== البداية ====================

class RootScreen extends StatefulWidget {
  const RootScreen({super.key});
  @override
  State<RootScreen> createState() => _RootScreenState();
}

class _RootScreenState extends State<RootScreen> {
  MoneyData? data;
  bool loading = true;

  @override
  void initState() {
    super.initState();
    _load();
  }

  Future<void> _load() async {
    final d = await Store.load();
    setState(() {
      data = d;
      loading = false;
    });
  }

  @override
  Widget build(BuildContext context) {
    if (loading) {
      return const Scaffold(body: Center(child: CircularProgressIndicator()));
    }
    if (data == null || data!.role.isEmpty) {
      return SetupScreen(onDone: (d) async {
        await Store.save(d);
        setState(() => data = d);
      });
    }
    return HomeShell(data: data!, onReplace: (d) async {
      await Store.save(d);
      setState(() => data = d);
    });
  }
}

/// أول تشغيل: التطبيق فاضي. يا ترجّع نسخة، يا تبدأ جديد وتختار التليفون ده إيه.
class SetupScreen extends StatefulWidget {
  final Future<void> Function(MoneyData) onDone;
  const SetupScreen({super.key, required this.onDone});
  @override
  State<SetupScreen> createState() => _SetupScreenState();
}

class _SetupScreenState extends State<SetupScreen> {
  String role = 'main'; // داخلي بس (للتوافق مع النسخ القديمة)
  String? nameErr;
  final nameCtl = TextEditingController();
  List<FoundFile>? backups;

  @override
  void initState() {
    super.initState();
    // لو الصلاحية موجودة أصلاً: ندوّر على النسخ لوحدنا ونعرضها
    () async {
      if (await hasSharedStorageAccess()) {
        final b = await listBackups();
        if (mounted && b.isNotEmpty) setState(() => backups = b);
      }
    }();
  }

  MoneyData? newData() {
    final n = cleanLabel(nameCtl.text);
    if (n.isEmpty) {
      setState(() => nameErr = 'اكتب اسم للتليفون الأول');
      return null;
    }
    final d = MoneyData.fresh();
    d.role = role;
    d.phoneName = n;
    return d;
  }

  Future<void> findBackups() async {
    await getAppFolder(requestPermissions: true);
    final b = await listBackups();
    setState(() => backups = b);
    if (b.isEmpty && mounted) toast(context, 'مفيش نسخ احتياطية لقيناها على الموبايل');
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: SafeArea(
        child: ListView(padding: const EdgeInsets.all(20), children: [
          const SizedBox(height: 10),
          const BrandHeader(),
          const SizedBox(height: 22),
          const Text('أهلاً! سمّي التليفون ده', style: TextStyle(fontSize: 20, fontWeight: FontWeight.w800)),
          const SizedBox(height: 6),
          const Text('كل تليفون ليه أرصدته لوحده، والتليفونين بيبعتوا نسخهم لبعض في أي وقت عشان يطلع إجمالي الصرف. '
              'الاسم ده هو اللي هيظهر على التليفون التاني.',
              style: TextStyle(color: C.muted)),
          const SizedBox(height: 8),
          const SizedBox(height: 6),
          TextField(
            controller: nameCtl,
            decoration: InputDecoration(labelText: 'اسم التليفون — مثلاً: تليفون أحمد', errorText: nameErr),
            onChanged: (_) {
              if (nameErr != null) setState(() => nameErr = null);
            },
          ),
          const SizedBox(height: 18),
          FilledButton.icon(
            style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 16)),
            onPressed: () async {
              final d = newData();
              if (d != null) await widget.onDone(d);
            },
            icon: const Icon(Icons.rocket_launch_rounded),
            label: const Text('ابدأ (فاضي)', style: TextStyle(fontSize: 16, fontWeight: FontWeight.w800)),
          ),
          ...[
            const SizedBox(height: 10),
            OutlinedButton.icon(
              style: OutlinedButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
              onPressed: () async {
                final d = newData();
                if (d == null) return;
                final changed = await receiveListsFlow(context, d);
                if (changed) await widget.onDone(d);
              },
              icon: const Icon(Icons.playlist_add_rounded),
              label: const Text('عندك ملف قوايم من التليفون التاني؟ ابدأ بيه'),
            ),
            const Padding(
              padding: EdgeInsets.only(top: 4),
              child: Text('اختياري: الأسرة والبنود والفواتير بتيجي جاهزة، والأرصدة بتبدأ صفر.',
                  style: TextStyle(color: C.muted, fontSize: 12)),
            ),
          ],
          const SizedBox(height: 10),
          OutlinedButton.icon(
            style: OutlinedButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
            onPressed: findBackups,
            icon: const Icon(Icons.restore_rounded),
            label: const Text('عندي نسخة احتياطية — رجّعها'),
          ),
          if (backups != null)
            for (final b in backups!)
              Padding(
                padding: const EdgeInsets.only(top: 8),
                child: ListTile(
                  tileColor: C.card,
                  shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14)),
                  leading: const Icon(Icons.backup_rounded, color: C.emerald),
                  title: Text(b.label, style: const TextStyle(fontSize: 13)),
                  onTap: () async {
                    final ok = await confirmDialog(context, 'ترجّع النسخة دي؟', b.label, yes: 'رجّعها');
                    if (!ok) return;
                    final d = MoneyData.fromJson(b.json);
                    if (d.role.isEmpty) d.role = 'main';
                    await widget.onDone(d);
                  },
                ),
              ),
          const SizedBox(height: 16),
          const Text('التطبيق بيحفظ نسخة تلقائي كل يوم في فولدر «A2R Money/Backups» (آخر ١٤ يوم). '
              'لما تمسح التطبيق وتنزله تاني دوس «رجّعها».',
              style: TextStyle(color: C.muted, fontSize: 12)),
        ]),
      ),
    );
  }
}

class BrandHeader extends StatelessWidget {
  final bool small;
  const BrandHeader({super.key, this.small = false});
  @override
  Widget build(BuildContext context) {
    final s = small ? 0.62 : 1.0;
    return Row(mainAxisSize: MainAxisSize.min, children: [
      Container(
        width: 46 * s,
        height: 46 * s,
        decoration: BoxDecoration(
          shape: BoxShape.circle,
          gradient: const LinearGradient(colors: [C.gold, C.goldDeep], begin: Alignment.topLeft, end: Alignment.bottomRight),
          border: Border.all(color: const Color(0xFFFFE39A), width: 2 * s),
        ),
        alignment: Alignment.center,
        child: Text('ج', style: TextStyle(fontSize: 24 * s, fontWeight: FontWeight.w900, color: const Color(0xFF3A2A00))),
      ),
      SizedBox(width: 10 * s),
      Column(crossAxisAlignment: CrossAxisAlignment.start, mainAxisSize: MainAxisSize.min, children: [
        Text('A2R', style: TextStyle(fontSize: 26 * s, height: 1, fontWeight: FontWeight.w900, color: C.gold, letterSpacing: 1)),
        Text('MONEY', style: TextStyle(fontSize: 12 * s, height: 1.2, fontWeight: FontWeight.w800, color: C.emerald, letterSpacing: 4 * s)),
      ]),
    ]);
  }
}

// ==================== الشاشة الرئيسية ====================

class HomeShell extends StatefulWidget {
  final MoneyData data;
  final Future<void> Function(MoneyData) onReplace;
  const HomeShell({super.key, required this.data, required this.onReplace});
  @override
  State<HomeShell> createState() => _HomeShellState();
}

class _HomeShellState extends State<HomeShell> with WidgetsBindingObserver {
  int tab = 0;
  String month = monthKeyOf(DateTime.now());
  Timer? _backupTimer;
  MoneyData get d => widget.data;

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(this);
    WidgetsBinding.instance.addPostFrameCallback((_) => _checkStorage());
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    _backupTimer?.cancel();
    super.dispose();
  }

  @override
  void didChangeAppLifecycleState(AppLifecycleState s) {
    if (s == AppLifecycleState.paused) {
      _backupTimer?.cancel();
      writeAutoBackup(d);
    }
  }

  Future<void> _checkStorage() async {
    if (await hasSharedStorageAccess()) {
      await writeAutoBackup(d);
      return;
    }
    if (!mounted) return;
    final ok = await confirmDialog(
        context,
        'صلاحية الملفات',
        'التطبيق محتاج صلاحية «كل الملفات» عشان يحفظ نسخة احتياطية تلقائي في فولدر A2R Money '
            '— فتفضل بياناتك لو مسحت التطبيق أو نزلت تحديث.',
        yes: 'اسمح',
        no: 'بعدين');
    if (ok) {
      await getAppFolder(requestPermissions: true);
      await writeAutoBackup(d);
    }
  }

  Future<void> persist() async {
    await Store.save(d);
    _backupTimer?.cancel();
    _backupTimer = Timer(const Duration(seconds: 4), () => writeAutoBackup(d));
    if (mounted) setState(() {});
  }

  Future<void> openForm(String type, {Txn? existing, Bill? bill, String? billMonth, String? personId}) async {
    final r = await Navigator.push<TxnFormResult>(
      context,
      MaterialPageRoute(
          builder: (_) => TxnFormScreen(
              data: d, type: type, existing: existing, bill: bill, billMonth: billMonth, presetPersonId: personId)),
    );
    if (r == null) {
      await persist(); // ممكن يكون اتضاف شخص أو بند
      return;
    }
    if (r.resend && existing != null) {
      await sendTransfer(existing);
      return;
    }
    if (r.deleted && existing != null) {
      d.removeTxn(existing.id);
      await persist();
      if (mounted) toast(context, 'اتمسحت');
      return;
    }
    if (r.saved != null) {
      d.upsert(r.saved!);
      await persist();
      if (mounted) {
        final t = r.saved!;
        toast(context,
            '${kTypeLabels[t.type]} ${egp(t.amount)} • ${d.placeName(t.placeId)} دلوقتي ${egp(d.balanceOf(t.placeId))}');
        if (t.type == tToPhone && existing == null) {
          final go = await confirmDialog(context, 'تبعت التحويل؟',
              'هيتعمل ملف صغير بـ ${egp(t.amount)} تبعته على واتساب للتليفون التاني.\n'
                  'وتقدر تبعته تاني بعدين من العملية نفسها.',
              yes: 'ابعته', no: 'بعدين');
          if (go) await sendTransfer(t);
        }
      }
    }
  }

  /// ملف التحويل ← واتساب (أو أي برنامج)
  Future<void> sendTransfer(Txn t) async {
    try {
      Directory out;
      try {
        final dir = await getAppFolder();
        out = Directory('${dir.path}/Transfers');
        if (!await out.exists()) await out.create(recursive: true);
      } catch (_) {
        out = await getTemporaryDirectory();
      }
      final f = File('${out.path}/$kTransferFilePrefix${t.date}_${t.id.substring(max(0, t.id.length - 8))}.json');
      await f.writeAsString(jsonEncode(d.transferPayload(t)), flush: true);
      await Share.shareXFiles([XFile(f.path, mimeType: 'application/json')],
          text: 'A2R Money — تحويل ${egp(t.amount)} للتليفون التاني');
    } catch (e) {
      if (mounted) toast(context, 'الملف ماتبعتش: $e', bad: true);
    }
  }

  /// استلام تحويل من التليفون التاني (بإيدك من أي فولدر)
  Future<void> receiveTransfer() async {
    await getAppFolder(requestPermissions: true);
    final files = await listTransferFiles(d);
    if (!mounted) return;
    final pick = await showModalBottomSheet<FoundFile>(
      context: context,
      isScrollControlled: true,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            Text('تحويلات من ${d.otherName}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            const Text('من واتساب و Download — افتح الملف في الشات مرة الأول عشان يتنزل', style: TextStyle(color: C.muted, fontSize: 12)),
            const SizedBox(height: 10),
            if (files.isEmpty)
              const Padding(
                padding: EdgeInsets.all(16),
                child: Text('مفيش ملفات تحويل. خلي التليفون التاني يسجّل «حوّل فلوس» ويبعت الملف على واتساب.',
                    style: TextStyle(color: C.muted)),
              ),
            Flexible(
              child: ListView(shrinkWrap: true, children: [
                for (final f in files)
                  ListTile(
                    leading: Icon(Icons.install_mobile_rounded,
                        color: d.hasReceivedTransfer('${f.json['id']}') ? C.muted : C.emerald),
                    title: Text(f.label, style: const TextStyle(fontWeight: FontWeight.w700)),
                    subtitle: Text(
                        [
                          'من ${f.json['fromPhoneName'] ?? d.otherName}',
                          if ('${f.json['note'] ?? ''}'.trim().isNotEmpty) '${f.json['note']}',
                        ].join(' • '),
                        style: const TextStyle(color: C.muted, fontSize: 12)),
                    onTap: () => Navigator.pop(c, f),
                  ),
              ]),
            ),
          ]),
        ),
      ),
    );
    if (pick == null || !mounted) return;
    if (d.hasReceivedTransfer('${pick.json['id']}')) {
      toast(context, 'التحويل ده اتسجل قبل كده', bad: true);
      return;
    }
    final amt = (pick.json['amount'] as num?)?.toInt() ?? 0;
    final placeId = await pickPlace(context, d, title: 'الـ ${egp(amt)} هتدخل فين؟');
    if (placeId == null || !mounted) return;
    final r = d.importTransferJson(pick.json, placeId);
    if (r.ok) await persist();
    if (mounted) toast(context, r.message, bad: !r.ok);
  }

  /// البادج اللي فوق ← تغيير اسم التليفون
  Future<void> phoneSheet() async {
    final n = await askText(context, 'اسم التليفون ده', initial: d.phoneName, hint: 'مثلاً: تليفون أحمد');
    if (n == null) return;
    d.phoneName = n;
    await persist();
  }

  /// «ابعت نسختي»: الشهر ده أو كل الشهور + القوايم ← واتساب
  Future<void> sendCopy() async {
    final scope = await showModalBottomSheet<String>(
      context: context,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            Text('ابعت نسختي لـ ${d.otherName}', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            const Text('عملياتك + القوايم. التليفون التاني بيبدّل نسختك القديمة بالجديدة — مفيش تكرار، وأرصدته مابتتلمسش.',
                style: TextStyle(color: C.muted, fontSize: 12.5)),
            const SizedBox(height: 12),
            ListTile(
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14), side: const BorderSide(color: C.line)),
              leading: const Icon(Icons.all_inclusive_rounded, color: C.gold),
              title: const Text('كل الشهور (أحسن)'),
              onTap: () => Navigator.pop(c, 'all'),
            ),
            const SizedBox(height: 8),
            ListTile(
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(14), side: const BorderSide(color: C.line)),
              leading: const Icon(Icons.calendar_month_rounded, color: C.emerald),
              title: Text('${monthLabel(month)} بس'),
              onTap: () => Navigator.pop(c, month),
            ),
          ]),
        ),
      ),
    );
    if (scope == null || !mounted) return;
    try {
      Directory out;
      try {
        final dir = await getAppFolder();
        out = Directory('${dir.path}/Copies');
        if (!await out.exists()) await out.create(recursive: true);
      } catch (_) {
        out = await getTemporaryDirectory();
      }
      final now = DateTime.now();
      final f = File('${out.path}/$kCopyFilePrefix${isoDate(now)}_${two(now.hour)}${two(now.minute)}${two(now.second)}.json');
      await f.writeAsString(jsonEncode(d.exportCopy(month: scope == 'all' ? null : scope)), flush: true);
      await Share.shareXFiles([XFile(f.path, mimeType: 'application/json')],
          text: 'A2R Money — نسخة ${d.badgeName}${scope == 'all' ? '' : ' (${monthLabel(scope)})'}');
    } catch (e) {
      if (mounted) toast(context, 'الملف ماتبعتش: $e', bad: true);
    }
  }

  /// اختيار فولدر الاستلام (بيتحفظ)
  Future<bool> chooseSyncFolder() async {
    await getAppFolder(requestPermissions: true);
    if (!mounted) return false;
    final p = await Navigator.push<String>(context, MaterialPageRoute(builder: (_) => FolderPickerScreen(current: d.syncFolder)));
    if (p == null || p.isEmpty) return false;
    d.syncFolder = p;
    await persist();
    return true;
  }

  /// «تحديث»: أحدث نسخة من التليفون التاني من الفولدر المحفوظ + القوايم + التحويلات
  Future<void> syncNow() async {
    if (d.syncFolder.isEmpty) {
      final ok = await chooseSyncFolder();
      if (!ok || !mounted) return;
    }
    await getAppFolder(requestPermissions: true);
    List<Map<String, dynamic>> files;
    try {
      if (!await Directory(d.syncFolder).exists()) {
        if (!mounted) return;
        final change = await confirmDialog(context, 'الفولدر مش موجود', '${d.syncFolder}\nتختار فولدر تاني؟', yes: 'اختار');
        if (change) await chooseSyncFolder();
        return;
      }
      files = await readSyncFolder(d.syncFolder);
    } catch (e) {
      if (mounted) toast(context, 'مقدرتش أقرا الفولدر: $e', bad: true);
      return;
    }
    final plan = d.planSync(files);
    if (!mounted) return;
    if (plan.isEmpty) {
      toast(context, 'مفيش نسخة أحدث من ${d.otherName}');
      return;
    }
    final msgs = <String>[];
    var changed = false;
    String latestName = '', latestAt = '';
    for (final j in [...plan.copies, ...plan.months]) {
      final r = d.importCopyJson(j);
      if (r.ok) {
        changed = true;
        if (r.exportedAt.compareTo(latestAt) > 0) {
          latestAt = r.exportedAt;
          latestName = r.fromName;
        }
      }
    }
    if (latestAt.isNotEmpty) {
      final w = DateTime.tryParse(latestAt)?.toLocal();
      msgs.add('اتحدث — آخر نسخة من $latestName ${w == null ? '' : '${arDateShort(isoDate(w))} ${two(w.hour)}:${two(w.minute)}'}');
    }
    var listsAdded = 0;
    for (final j in plan.lists) {
      final r = d.mergeListsJson(j, addPlaces: false);
      if (r.ok && r.total > 0) {
        listsAdded += r.total;
        changed = true;
      }
    }
    if (listsAdded > 0) msgs.add('اتضاف $listsAdded من القوايم');
    if (changed) await persist();
    if (!mounted) return;
    // التحويلات: كل واحد بيختار المكان اللي هيدخله (ولو اتلغى يفضل للمرة الجاية)
    var trDone = 0;
    for (final j in plan.transfers) {
      if (!mounted) return;
      final amt = (j['amount'] as num?)?.toInt() ?? 0;
      final pid = await pickPlace(context, d, title: 'تحويل ${egp(amt)} من ${j['fromPhoneName'] ?? d.otherName} — هيدخل فين؟');
      if (pid == null) continue;
      final r = d.importTransferJson(j, pid);
      if (r.ok) {
        trDone++;
        await persist();
      }
    }
    if (trDone > 0) msgs.add('اتسجل $trDone تحويل');
    if (plan.transfers.length > trDone) msgs.add('${plan.transfers.length - trDone} تحويل لسه مستني (دوس تحديث تاني)');
    if (mounted) toast(context, msgs.isEmpty ? 'مفيش نسخة أحدث من ${d.otherName}' : msgs.join(' • '));
  }

  /// «ابدأ بأرصدتك»: كل الأماكن في شاشة واحدة
  Future<void> openStartBalances() async {
    final ok = await Navigator.push<bool>(context, MaterialPageRoute(builder: (_) => StartBalancesScreen(data: d)));
    if (ok == true) {
      await persist();
      if (mounted) toast(context, 'اتحفظت الأرصدة • الإجمالي ${egp(d.totalBalance)}');
    }
  }

  Future<void> openTxn(Txn t) => openForm(t.type, existing: t, bill: t.type == tBill ? d.billById(t.billId) : null);

  Future<void> deleteTxn(Txn t) async {
    final ok = await confirmDialog(context, 'مسح العملية؟', '${kTypeLabels[t.type]} ${egp(t.amount)} • ${arDateShort(t.date)}',
        yes: 'امسح', danger: true);
    if (!ok) return;
    d.removeTxn(t.id);
    await persist();
  }

  void setMonth(String m) => setState(() => month = m);
  void goTab(int i) => setState(() => tab = i);

  void quickAdd() {
    showModalBottomSheet(
      context: context,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 18, 16, 14),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            const Text('تسجّل إيه؟', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            const SizedBox(height: 12),
            GridView.count(
              crossAxisCount: 3,
              shrinkWrap: true,
              mainAxisSpacing: 10,
              crossAxisSpacing: 10,
              physics: const NeverScrollableScrollPhysics(),
              children: [
                for (final t in [tExpense, tWithdraw, tAdd, tTransfer, tAdvance, tRepay, tToPhone, tFromPhone])
                  InkWell(
                    borderRadius: BorderRadius.circular(18),
                    onTap: () {
                      Navigator.pop(c);
                      if (t == tFromPhone) {
                        receiveTransfer();
                      } else {
                        openForm(t);
                      }
                    },
                    child: Container(
                      decoration: BoxDecoration(
                          color: typeColor(t).withValues(alpha: 0.12),
                          borderRadius: BorderRadius.circular(18),
                          border: Border.all(color: typeColor(t).withValues(alpha: 0.45))),
                      child: Column(mainAxisAlignment: MainAxisAlignment.center, children: [
                        Icon(typeIcon(t), color: typeColor(t), size: 30),
                        const SizedBox(height: 6),
                        Text(t == tToPhone ? 'حوّل لـ ${d.otherName}' : t == tFromPhone ? 'استلم تحويل' : kTypeLabels[t]!, maxLines: 2, overflow: TextOverflow.ellipsis,
                            textAlign: TextAlign.center, style: const TextStyle(fontWeight: FontWeight.w800)),
                      ]),
                    ),
                  ),
              ],
            ),
          ]),
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final due = d.dueBills(monthKeyOf(DateTime.now()), DateTime.now()).length;
    final pages = [
      HomeTab(shell: this),
      RecordsTab(shell: this),
      AdvancesTab(shell: this),
      BillsTab(shell: this),
      MonthTab(shell: this),
    ];
    return Scaffold(
      appBar: AppBar(
        titleSpacing: 16,
        title: const BrandHeader(small: true),
        actions: [
          Center(child: PhoneBadge(data: d, onTap: phoneSheet)),
          IconButton(
            tooltip: 'الإعدادات',
            icon: const Icon(Icons.tune_rounded),
            onPressed: () async {
              await Navigator.push(context, MaterialPageRoute(builder: (_) => SettingsScreen(shell: this)));
              await persist();
            },
          ),
        ],
      ),
      body: pages[tab],
      floatingActionButton: tab <= 2
          ? FloatingActionButton.extended(
              onPressed: tab == 2 ? () => openForm(tAdvance) : quickAdd,
              icon: const Icon(Icons.add_rounded),
              label: Text(tab == 2 ? 'سلفة' : 'تسجيل', style: const TextStyle(fontWeight: FontWeight.w800)),
            )
          : null,
      bottomNavigationBar: NavigationBar(
        selectedIndex: tab,
        onDestinationSelected: (i) => setState(() => tab = i),
        destinations: [
          const NavigationDestination(icon: Icon(Icons.home_outlined), selectedIcon: Icon(Icons.home_rounded), label: 'الرئيسية'),
          const NavigationDestination(icon: Icon(Icons.list_alt_outlined), selectedIcon: Icon(Icons.list_alt_rounded), label: 'العمليات'),
          const NavigationDestination(icon: Icon(Icons.handshake_outlined), selectedIcon: Icon(Icons.handshake_rounded), label: 'السلف'),
          NavigationDestination(
              icon: Badge(isLabelVisible: due > 0, label: Text('$due'), child: const Icon(Icons.receipt_long_outlined)),
              selectedIcon: Badge(isLabelVisible: due > 0, label: Text('$due'), child: const Icon(Icons.receipt_long_rounded)),
              label: 'الفواتير'),
          const NavigationDestination(icon: Icon(Icons.insights_outlined), selectedIcon: Icon(Icons.insights_rounded), label: 'الشهر'),
        ],
      ),
    );
  }
}

// ---------- الرئيسية ----------

class HomeTab extends StatelessWidget {
  final _HomeShellState shell;
  const HomeTab({super.key, required this.shell});

  @override
  Widget build(BuildContext context) {
    final d = shell.d;
    final mk = monthKeyOf(DateTime.now());
    final s = d.ownSummary(mk);
    final total = d.totalBalance;
    final due = d.dueBills(mk, DateTime.now());
    final recent = d.monthTxns(mk).take(5).toList();
    return ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 100), children: [
      if (d.isEmptyStart)
        Padding(
          padding: const EdgeInsets.only(bottom: 12),
          child: InkWell(
            borderRadius: BorderRadius.circular(22),
            onTap: shell.openStartBalances,
            child: Container(
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                borderRadius: BorderRadius.circular(22),
                color: C.gold.withValues(alpha: 0.12),
                border: Border.all(color: C.gold, width: 1.4),
              ),
              child: const Row(children: [
                Icon(Icons.account_balance_wallet_rounded, color: C.gold, size: 34),
                SizedBox(width: 12),
                Expanded(
                  child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                    Text('ابدأ بأرصدتك', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w900, color: C.gold)),
                    SizedBox(height: 2),
                    Text('اكتب الفلوس اللي معاك دلوقتي في كل مكان (كاش، بنك، محفظة) مرة واحدة',
                        style: TextStyle(fontSize: 12.5, color: C.muted)),
                  ]),
                ),
                Icon(Icons.chevron_left_rounded, color: C.gold),
              ]),
            ),
          ),
        ),
      Container(
        padding: const EdgeInsets.all(20),
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(26),
          gradient: const LinearGradient(
              colors: [Color(0xFF0E5A50), Color(0xFF0A3D3A), Color(0xFF173B2A)],
              begin: Alignment.topRight,
              end: Alignment.bottomLeft),
          border: Border.all(color: C.gold.withValues(alpha: 0.45)),
          boxShadow: [BoxShadow(color: Colors.black.withValues(alpha: 0.35), blurRadius: 18, offset: const Offset(0, 8))],
        ),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            const Text('إجمالي الرصيد', style: TextStyle(color: C.muted, fontWeight: FontWeight.w700)),
            const Spacer(),
            Text(d.phoneName, style: const TextStyle(color: C.muted, fontSize: 12)),
          ]),
          const SizedBox(height: 6),
          FittedBox(
            fit: BoxFit.scaleDown,
            alignment: AlignmentDirectional.centerStart,
            child: Text(egp(total),
                style: TextStyle(fontSize: 34, fontWeight: FontWeight.w900, color: total < 0 ? C.red : C.gold)),
          ),
          const SizedBox(height: 12),
          Row(children: [
            Expanded(child: _mini('صرف ${arabicMonths[DateTime.now().month - 1]}', egp(s.totalSpend), C.red)),
            const SizedBox(width: 10),
            Expanded(child: _mini('اتضاف الشهر ده', egp(s.added), C.emerald)),
          ]),
        ]),
      ),
      const SectionTitle('الأماكن'),
      SizedBox(
        height: 118,
        child: ListView(scrollDirection: Axis.horizontal, children: [
          for (final p in d.activePlaces)
            Padding(
              padding: const EdgeInsetsDirectional.only(end: 10),
              child: _placeCard(context, d, p),
            ),
        ]),
      ),
      const SizedBox(height: 12),
      Row(children: [
        for (final t in [tExpense, tWithdraw, tAdd, tTransfer])
          Expanded(
            child: Padding(
              padding: const EdgeInsets.symmetric(horizontal: 4),
              child: InkWell(
                borderRadius: BorderRadius.circular(16),
                onTap: () => shell.openForm(t),
                child: Container(
                  padding: const EdgeInsets.symmetric(vertical: 12),
                  decoration: BoxDecoration(
                      color: C.card, borderRadius: BorderRadius.circular(16), border: Border.all(color: C.line)),
                  child: Column(children: [
                    Icon(typeIcon(t), color: typeColor(t)),
                    const SizedBox(height: 4),
                    Text(kTypeLabels[t]!, style: const TextStyle(fontSize: 12, fontWeight: FontWeight.w700)),
                  ]),
                ),
              ),
            ),
          ),
      ]),
      const SizedBox(height: 10),
      Row(children: [
        Expanded(
          child: FilledButton.icon(
            style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 13)),
            onPressed: shell.syncNow,
            icon: const Icon(Icons.sync_rounded),
            label: const Text('تحديث', style: TextStyle(fontWeight: FontWeight.w800)),
          ),
        ),
        const SizedBox(width: 8),
        Expanded(
          child: OutlinedButton.icon(
            style: OutlinedButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 13)),
            onPressed: shell.sendCopy,
            icon: const Icon(Icons.ios_share_rounded),
            label: Text('ابعت نسختي لـ ${d.otherName}', maxLines: 1, overflow: TextOverflow.ellipsis,
                style: const TextStyle(fontWeight: FontWeight.w700)),
          ),
        ),
      ]),
      const SizedBox(height: 8),
      OutlinedButton.icon(
        style: OutlinedButton.styleFrom(
            padding: const EdgeInsets.symmetric(vertical: 12), side: BorderSide(color: C.blue.withValues(alpha: 0.7))),
        onPressed: () => shell.openForm(tToPhone),
        icon: const Icon(Icons.send_to_mobile_rounded, color: C.blue),
        label: Text('حوّل فلوس لـ ${d.otherName}', style: const TextStyle(fontWeight: FontWeight.w800)),
      ),
      if (due.isNotEmpty) ...[
        const SectionTitle('فواتير لازم تتدفع'),
        for (final b in due)
          Padding(
            padding: const EdgeInsets.only(bottom: 8),
            child: ListTile(
              tileColor: C.red.withValues(alpha: 0.10),
              shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(16), side: BorderSide(color: C.red.withValues(alpha: 0.5))),
              leading: const Icon(Icons.notifications_active_rounded, color: C.red),
              title: Text(b.name, style: const TextStyle(fontWeight: FontWeight.w800)),
              subtitle: Text(billDueText(b, mk), style: const TextStyle(color: C.muted)),
              trailing: FilledButton(
                  onPressed: () => shell.openForm(tBill, bill: b, billMonth: mk), child: const Text('ادفع')),
            ),
          ),
      ],
      if (d.totalAdvancesOut > 0) ...[
        const SectionTitle('السلف'),
        ListTile(
          tileColor: C.card,
          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16), side: const BorderSide(color: C.line)),
          leading: const Icon(Icons.handshake_rounded, color: C.violet),
          title: const Text('فلوس عند الناس', style: TextStyle(fontWeight: FontWeight.w700)),
          trailing: Text(egp(d.totalAdvancesOut), style: const TextStyle(fontWeight: FontWeight.w900, color: C.violet)),
          onTap: () => shell.goTab(2),
        ),
      ],
      SectionTitle('آخر العمليات',
          trailing: TextButton(onPressed: () => shell.goTab(1), child: const Text('الكل'))),
      if (recent.isEmpty)
        const Padding(
          padding: EdgeInsets.all(20),
          child: Text('لسه مفيش عمليات الشهر ده — دوس «تسجيل» تحت', textAlign: TextAlign.center, style: TextStyle(color: C.muted)),
        ),
      for (final t in recent) TxnTile(data: d, t: t, onTap: () => shell.openTxn(t), onLongPress: () => shell.deleteTxn(t)),
    ]);
  }

  Widget _mini(String l, String v, Color c) => Container(
        padding: const EdgeInsets.all(10),
        decoration: BoxDecoration(color: Colors.black.withValues(alpha: 0.18), borderRadius: BorderRadius.circular(14)),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Text(l, style: const TextStyle(color: C.muted, fontSize: 11.5)),
          FittedBox(
              fit: BoxFit.scaleDown,
              alignment: AlignmentDirectional.centerStart,
              child: Text(v, style: TextStyle(fontWeight: FontWeight.w900, color: c, fontSize: 16))),
        ]),
      );

  Widget _placeCard(BuildContext context, MoneyData d, Place p) {
    final b = d.balanceOf(p.id);
    final c = placeColor(p.kind);
    return InkWell(
      borderRadius: BorderRadius.circular(20),
      onTap: () => Navigator.push(context, MaterialPageRoute(builder: (_) => PlaceDetailsScreen(shell: shell, place: p))),
      child: Container(
        width: 150,
        padding: const EdgeInsets.all(14),
        decoration: BoxDecoration(
          color: C.card,
          borderRadius: BorderRadius.circular(20),
          border: Border.all(color: c.withValues(alpha: 0.55)),
          gradient: LinearGradient(colors: [c.withValues(alpha: 0.16), C.card], begin: Alignment.topCenter, end: Alignment.bottomCenter),
        ),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Row(children: [
            Icon(placeIcon(p.kind), color: c, size: 22),
            const SizedBox(width: 6),
            Expanded(child: Text(p.name, style: const TextStyle(fontWeight: FontWeight.w800), overflow: TextOverflow.ellipsis)),
          ]),
          const Spacer(),
          FittedBox(
            fit: BoxFit.scaleDown,
            alignment: AlignmentDirectional.centerStart,
            child: Text(money(b), style: TextStyle(fontSize: 22, fontWeight: FontWeight.w900, color: b < 0 ? C.red : C.text)),
          ),
          const Text('جنيه', style: TextStyle(color: C.muted, fontSize: 11)),
        ]),
      ),
    );
  }
}

String billDueText(Bill b, String mk) {
  final due = b.dueIn(mk);
  final now = DateTime.now();
  final diff = due.difference(DateTime(now.year, now.month, now.day)).inDays;
  final when = '${arabicWeekdays[due.weekday - 1]} ${due.day}/${due.month}';
  if (diff < 0) return 'متأخرة ${-diff} يوم (معادها $when)';
  if (diff == 0) return 'معادها النهارده';
  return 'معادها $when (فاضل $diff يوم)';
}

class PlaceDetailsScreen extends StatefulWidget {
  final _HomeShellState shell;
  final Place place;
  const PlaceDetailsScreen({super.key, required this.shell, required this.place});
  @override
  State<PlaceDetailsScreen> createState() => _PlaceDetailsScreenState();
}

class _PlaceDetailsScreenState extends State<PlaceDetailsScreen> {
  @override
  Widget build(BuildContext context) {
    final d = widget.shell.d;
    final p = widget.place;
    final ts = d.txns.where((t) => t.placeId == p.id || t.toPlaceId == p.id).toList()
      ..sort((a, b) => b.date.compareTo(a.date) != 0 ? b.date.compareTo(a.date) : b.createdAt.compareTo(a.createdAt));
    final b = d.balanceOf(p.id);
    return Scaffold(
      appBar: AppBar(title: Text(p.name)),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 40), children: [
        Container(
          padding: const EdgeInsets.all(18),
          decoration: BoxDecoration(
              color: C.card, borderRadius: BorderRadius.circular(22), border: Border.all(color: placeColor(p.kind))),
          child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            const Text('الرصيد', style: TextStyle(color: C.muted)),
            Text(egp(b), style: TextStyle(fontSize: 30, fontWeight: FontWeight.w900, color: b < 0 ? C.red : C.gold)),
            Text('رصيد البداية: ${egp(p.opening)}', style: const TextStyle(color: C.muted, fontSize: 12)),
          ]),
        ),
        const SizedBox(height: 12),
        Row(children: [
          Expanded(
            child: FilledButton.icon(
              onPressed: () async {
                widget.shell.d.lastInPlaceId = p.id;
                await widget.shell.openForm(tAdd);
                setState(() {});
              },
              icon: const Icon(Icons.add_card_rounded),
              label: const Text('إضافة رصيد'),
            ),
          ),
          const SizedBox(width: 10),
          Expanded(
            child: FilledButton.tonalIcon(
              onPressed: () async {
                widget.shell.d.lastPlaceId = p.id;
                await widget.shell.openForm(tWithdraw);
                setState(() {});
              },
              icon: const Icon(Icons.outbox_rounded),
              label: const Text('سحب'),
            ),
          ),
        ]),
        const SectionTitle('حركة المكان'),
        if (ts.isEmpty) const Text('مفيش حركة لسه', style: TextStyle(color: C.muted)),
        for (final t in ts)
          TxnTile(
              data: d,
              t: t,
              onTap: () async {
                widget.shell.openTxn(t);
              }),
      ]),
    );
  }
}

// ---------- العمليات (بحث وفلترة) ----------

class RecordsTab extends StatefulWidget {
  final _HomeShellState shell;
  const RecordsTab({super.key, required this.shell});
  @override
  State<RecordsTab> createState() => _RecordsTabState();
}

class _RecordsTabState extends State<RecordsTab> {
  String q = '';
  String filter = ''; // '' = الكل
  String placeFilter = '';

  @override
  Widget build(BuildContext context) {
    final sh = widget.shell;
    final d = sh.d;
    final nq = normalizeAr(q);
    final list = d.monthTxns(sh.month).where((t) {
      if (filter.isNotEmpty && t.type != filter) return false;
      if (placeFilter.isNotEmpty && t.placeId != placeFilter && t.toPlaceId != placeFilter) return false;
      if (nq.isEmpty) return true;
      final hay = normalizeAr([
        t.note,
        t.category,
        t.group,
        kTypeLabels[t.type] ?? '',
        d.personName(t.personId),
        d.placeName(t.placeId),
        if (t.toPlaceId.isNotEmpty) d.placeName(t.toPlaceId),
        money(t.amount).replaceAll(',', ''),
      ].join(' '));
      return hay.contains(nq);
    }).toList();
    final total = list.where((t) => t.isSpend).fold<int>(0, (s, t) => s + t.amount);
    String? lastDay;
    final children = <Widget>[];
    for (final t in list) {
      if (t.date != lastDay) {
        lastDay = t.date;
        children.add(Padding(
          padding: const EdgeInsets.fromLTRB(6, 10, 6, 6),
          child: Text(arDateShort(t.date), style: const TextStyle(color: C.gold, fontWeight: FontWeight.w800)),
        ));
      }
      children.add(TxnTile(data: d, t: t, onTap: () => sh.openTxn(t), onLongPress: () => sh.deleteTxn(t)));
    }
    return ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 100), children: [
      MonthSwitcher(month: sh.month, onChanged: (m) => sh.setMonth(m)),
      const SizedBox(height: 6),
      TextField(
        decoration: const InputDecoration(prefixIcon: Icon(Icons.search_rounded), hintText: 'دوّر: بند، شخص، ملاحظة، مبلغ…'),
        onChanged: (v) => setState(() => q = v),
      ),
      const SizedBox(height: 10),
      SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: Row(children: [
          for (final f in ['', tExpense, tBill, tWithdraw, tAdd, tTransfer, tAdvance, tRepay, tToPhone, tFromPhone])
            Padding(
              padding: const EdgeInsetsDirectional.only(end: 6),
              child: ChoiceChip(
                label: Text(f.isEmpty ? 'الكل' : kTypeLabels[f]!),
                selected: filter == f,
                showCheckmark: false,
                selectedColor: C.gold.withValues(alpha: 0.28),
                onSelected: (_) => setState(() => filter = f),
              ),
            ),
        ]),
      ),
      const SizedBox(height: 6),
      SingleChildScrollView(
        scrollDirection: Axis.horizontal,
        child: Row(children: [
          for (final p in [null, ...d.places])
            Padding(
              padding: const EdgeInsetsDirectional.only(end: 6),
              child: ChoiceChip(
                avatar: p == null ? null : Icon(placeIcon(p.kind), size: 16, color: placeColor(p.kind)),
                label: Text(p == null ? 'كل الأماكن' : p.name),
                selected: placeFilter == (p?.id ?? ''),
                showCheckmark: false,
                selectedColor: C.emerald.withValues(alpha: 0.25),
                onSelected: (_) => setState(() => placeFilter = p?.id ?? ''),
              ),
            ),
        ]),
      ),
      Padding(
        padding: const EdgeInsets.symmetric(vertical: 8, horizontal: 4),
        child: Text('${list.length} عملية • صرف ${egp(total)}', style: const TextStyle(color: C.muted)),
      ),
      if (list.isEmpty)
        const Padding(
            padding: EdgeInsets.all(24),
            child: Text('مفيش عمليات', textAlign: TextAlign.center, style: TextStyle(color: C.muted))),
      ...children,
      const Padding(
        padding: EdgeInsets.only(top: 8),
        child: Text('دوس على أي عملية تعدّلها، أو دوسة طويلة تمسحها', textAlign: TextAlign.center, style: TextStyle(color: C.muted, fontSize: 12)),
      ),
    ]);
  }
}

// ---------- السلف ----------

class AdvancesTab extends StatelessWidget {
  final _HomeShellState shell;
  const AdvancesTab({super.key, required this.shell});
  @override
  Widget build(BuildContext context) {
    final d = shell.d;
    final people = d.advancePeople
      ..sort((a, b) => d.advanceRemaining(b.id).compareTo(d.advanceRemaining(a.id)));
    return ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 100), children: [
      Container(
        padding: const EdgeInsets.all(18),
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(22),
          gradient: LinearGradient(colors: [C.violet.withValues(alpha: 0.25), C.card]),
          border: Border.all(color: C.violet.withValues(alpha: 0.5)),
        ),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          const Text('فلوس عند الناس (الباقي من السلف)', style: TextStyle(color: C.muted)),
          Text(egp(d.totalAdvancesOut), style: const TextStyle(fontSize: 28, fontWeight: FontWeight.w900, color: C.violet)),
        ]),
      ),
      const SectionTitle('الناس'),
      if (people.isEmpty)
        const Padding(
            padding: EdgeInsets.all(20),
            child: Text('مفيش سلف — دوس «سلفة» تحت', textAlign: TextAlign.center, style: TextStyle(color: C.muted))),
      for (final p in people)
        Padding(
          padding: const EdgeInsets.only(bottom: 8),
          child: ListTile(
            tileColor: C.card,
            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18), side: const BorderSide(color: C.line)),
            leading: CircleAvatar(
                backgroundColor: C.violet.withValues(alpha: 0.2),
                child: Text(p.name.isEmpty ? '?' : p.name.characters.first, style: const TextStyle(color: C.violet))),
            title: Text(p.name, style: const TextStyle(fontWeight: FontWeight.w800)),
            subtitle: Text(d.advanceRemaining(p.id) <= 0 ? 'خلّص' : 'الباقي عليه', style: const TextStyle(color: C.muted)),
            trailing: Text(egp(d.advanceRemaining(p.id)),
                style: TextStyle(
                    fontWeight: FontWeight.w900, color: d.advanceRemaining(p.id) <= 0 ? C.emerald : C.violet)),
            onTap: () => Navigator.push(context, MaterialPageRoute(builder: (_) => PersonAdvancesScreen(shell: shell, person: p))),
          ),
        ),
    ]);
  }
}

class PersonAdvancesScreen extends StatefulWidget {
  final _HomeShellState shell;
  final Person person;
  const PersonAdvancesScreen({super.key, required this.shell, required this.person});
  @override
  State<PersonAdvancesScreen> createState() => _PersonAdvancesScreenState();
}

class _PersonAdvancesScreenState extends State<PersonAdvancesScreen> {
  @override
  Widget build(BuildContext context) {
    final d = widget.shell.d;
    final p = widget.person;
    final ts = d.txns.where((t) => t.personId == p.id && (t.type == tAdvance || t.type == tRepay)).toList()
      ..sort((a, b) => b.date.compareTo(a.date));
    final given = ts.where((t) => t.type == tAdvance).fold<int>(0, (s, t) => s + t.amount);
    final back = ts.where((t) => t.type == tRepay).fold<int>(0, (s, t) => s + t.amount);
    final rem = d.advanceRemaining(p.id);
    return Scaffold(
      appBar: AppBar(title: Text('سلف ${p.name}')),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 40), children: [
        Row(children: [
          Expanded(child: _box('اتسلّف', egp(given), C.violet)),
          const SizedBox(width: 8),
          Expanded(child: _box('رجّع', egp(back), C.emerald)),
          const SizedBox(width: 8),
          Expanded(child: _box('الباقي', egp(rem), rem > 0 ? C.gold : C.emerald)),
        ]),
        const SizedBox(height: 12),
        Row(children: [
          Expanded(
            child: FilledButton.icon(
              onPressed: () async {
                await widget.shell.openForm(tRepay, personId: p.id);
                setState(() {});
              },
              icon: const Icon(Icons.keyboard_return_rounded),
              label: const Text('رد سلفة'),
            ),
          ),
          const SizedBox(width: 10),
          Expanded(
            child: FilledButton.tonalIcon(
              onPressed: () async {
                await widget.shell.openForm(tAdvance, personId: p.id);
                setState(() {});
              },
              icon: const Icon(Icons.handshake_rounded),
              label: const Text('سلفة تانية'),
            ),
          ),
        ]),
        const SectionTitle('السجل'),
        for (final t in ts)
          TxnTile(
              data: d,
              t: t,
              onTap: () async {
                await widget.shell.openTxn(t);
                setState(() {});
              }),
      ]),
    );
  }

  Widget _box(String l, String v, Color c) => Container(
        padding: const EdgeInsets.all(10),
        decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(16), border: Border.all(color: C.line)),
        child: Column(children: [
          Text(l, style: const TextStyle(color: C.muted, fontSize: 12)),
          FittedBox(fit: BoxFit.scaleDown, child: Text(v, style: TextStyle(fontWeight: FontWeight.w900, color: c))),
        ]),
      );
}

// ---------- الفواتير ----------

class BillsTab extends StatelessWidget {
  final _HomeShellState shell;
  const BillsTab({super.key, required this.shell});

  Future<void> editBill(BuildContext context, {Bill? b}) async {
    final d = shell.d;
    final r = await showDialog<Bill>(context: context, builder: (_) => BillDialog(bill: b, names: List.of(d.billNames)));
    if (r == null) return;
    r.name = d.addBillName(r.name);
    final i = d.bills.indexWhere((x) => x.id == r.id);
    if (i >= 0) {
      d.bills[i] = r;
    } else {
      d.bills.add(r);
    }
    await shell.persist();
  }

  @override
  Widget build(BuildContext context) {
    final d = shell.d;
    final mk = shell.month;
    final active = d.bills.where((b) => b.active).toList()..sort((a, b) => a.dueDay.compareTo(b.dueDay));
    final paid = active.where((b) => d.billPayment(b.id, mk) != null).length;
    final paidSum = active.fold<int>(0, (s, b) => s + (d.billPayment(b.id, mk)?.amount ?? 0));
    Widget group(String kind) {
      final l = active.where((b) => b.kind == kind).toList();
      if (l.isEmpty) return const SizedBox.shrink();
      return Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
        SectionTitle(kind == 'نت' ? 'فواتير النت' : 'فواتير البيت'),
        for (final b in l) _billTile(context, d, b, mk),
      ]);
    }

    return ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 100), children: [
      MonthSwitcher(month: mk, onChanged: shell.setMonth),
      Container(
        padding: const EdgeInsets.all(16),
        decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(20), border: Border.all(color: C.line)),
        child: Row(children: [
          const Icon(Icons.receipt_long_rounded, color: C.gold, size: 34),
          const SizedBox(width: 12),
          Expanded(
              child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            Text('اتدفع $paid من ${active.length}', style: const TextStyle(fontWeight: FontWeight.w800, fontSize: 16)),
            Text('مدفوع الشهر ده: ${egp(paidSum)}', style: const TextStyle(color: C.muted)),
          ])),
        ]),
      ),
      group('بيت'),
      group('نت'),
      if (active.isEmpty)
        const Padding(
            padding: EdgeInsets.all(20),
            child: Text('ضيف فواتيرك الشهرية (كهربا، ميه، غاز، نت…) — بتتكرر كل شهر لوحدها',
                textAlign: TextAlign.center, style: TextStyle(color: C.muted))),
      const SizedBox(height: 12),
      OutlinedButton.icon(
          onPressed: () => editBill(context),
          icon: const Icon(Icons.add_rounded),
          label: const Text('فاتورة جديدة')),
      if (d.bills.any((b) => !b.active)) ...[
        const SectionTitle('فواتير موقوفة'),
        for (final b in d.bills.where((b) => !b.active))
          ListTile(
            title: Text(b.name, style: const TextStyle(color: C.muted)),
            trailing: TextButton(onPressed: () => editBill(context, b: b), child: const Text('تعديل')),
          ),
      ],
    ]);
  }

  Widget _billTile(BuildContext context, MoneyData d, Bill b, String mk) {
    final pay = d.billPayment(b.id, mk);
    final now = DateTime.now();
    final late = pay == null && b.dueIn(mk).isBefore(DateTime(now.year, now.month, now.day));
    final c = pay != null ? C.emerald : (late ? C.red : C.gold);
    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Container(
        padding: const EdgeInsets.all(12),
        decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: c.withValues(alpha: 0.6))),
        child: Row(children: [
          Icon(b.kind == 'نت' ? Icons.wifi_rounded : Icons.home_work_rounded, color: c),
          const SizedBox(width: 10),
          Expanded(
            child: InkWell(
              onTap: () => editBill(context, b: b),
              child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
                Text(b.name, style: const TextStyle(fontWeight: FontWeight.w800, fontSize: 15)),
                Text(
                    pay != null
                        ? 'اتدفعت ${egp(pay.amount)} من ${d.placeName(pay.placeId)} • ${arDateShort(pay.date)}'
                        : billDueText(b, mk),
                    style: TextStyle(color: pay != null ? C.muted : c, fontSize: 12)),
              ]),
            ),
          ),
          if (pay == null)
            FilledButton(onPressed: () => shell.openForm(tBill, bill: b, billMonth: mk), child: const Text('ادفع'))
          else
            IconButton(
                tooltip: 'تعديل الدفع',
                onPressed: () => shell.openForm(tBill, existing: pay, bill: b, billMonth: mk),
                icon: const Icon(Icons.check_circle_rounded, color: C.emerald)),
        ]),
      ),
    );
  }
}

class BillDialog extends StatefulWidget {
  final Bill? bill;
  final List<String> names;
  const BillDialog({super.key, this.bill, this.names = kDefaultBillNames});
  @override
  State<BillDialog> createState() => _BillDialogState();
}

class _BillDialogState extends State<BillDialog> {
  late final TextEditingController name = TextEditingController(text: widget.bill?.name ?? '');
  late final TextEditingController usual =
      TextEditingController(text: (widget.bill?.usual ?? 0) > 0 ? money(widget.bill!.usual).replaceAll(',', '') : '');
  late String kind = widget.bill?.kind ?? 'بيت';
  late int day = widget.bill?.dueDay ?? 1;
  late bool active = widget.bill?.active ?? true;

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      title: Text(widget.bill == null ? 'فاتورة جديدة' : 'تعديل الفاتورة', style: const TextStyle(fontWeight: FontWeight.w800)),
      content: SingleChildScrollView(
        child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
          Wrap(spacing: 6, children: [
            for (final s in widget.names)
              ActionChip(label: Text(s), onPressed: () => setState(() {
                    name.text = s;
                    if (s.contains('نت') || s.contains('باقة')) kind = 'نت';
                  })),
          ]),
          const SizedBox(height: 8),
          TextField(
              controller: name,
              decoration: const InputDecoration(
                  labelText: 'اسم الفاتورة (اختار أو اكتب)', helperText: 'أي اسم جديد بيتحفظ ويطلعلك في الاختيارات')),
          const SizedBox(height: 10),
          SegmentedButton<String>(
            segments: const [
              ButtonSegment(value: 'بيت', label: Text('بيت'), icon: Icon(Icons.home_work_rounded)),
              ButtonSegment(value: 'نت', label: Text('نت'), icon: Icon(Icons.wifi_rounded)),
            ],
            selected: {kind},
            onSelectionChanged: (s) => setState(() => kind = s.first),
          ),
          const SizedBox(height: 10),
          Row(children: [
            const Text('يوم الدفع في الشهر:'),
            const Spacer(),
            IconButton(onPressed: () => setState(() => day = day <= 1 ? 31 : day - 1), icon: const Icon(Icons.remove_rounded)),
            Text('$day', style: const TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            IconButton(onPressed: () => setState(() => day = day >= 31 ? 1 : day + 1), icon: const Icon(Icons.add_rounded)),
          ]),
          TextField(
            controller: usual,
            keyboardType: const TextInputType.numberWithOptions(decimal: true),
            decoration: const InputDecoration(labelText: 'المبلغ المعتاد (اختياري)', suffixText: 'جنيه'),
          ),
          if (widget.bill != null)
            SwitchListTile(
              contentPadding: EdgeInsets.zero,
              title: const Text('شغالة كل شهر'),
              value: active,
              onChanged: (v) => setState(() => active = v),
            ),
        ]),
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(context), child: const Text('إلغاء')),
        FilledButton(
          onPressed: () {
            if (cleanLabel(name.text).isEmpty) return;
            Navigator.pop(
                context,
                Bill(
                  id: widget.bill?.id ?? newId('bi'),
                  name: cleanLabel(name.text),
                  kind: kind,
                  dueDay: day,
                  usual: parseMoney(usual.text) ?? 0,
                  active: active,
                ));
          },
          child: const Text('حفظ'),
        ),
      ],
    );
  }
}

// ---------- الشهر (الإجمالي + التليفون التاني + Excel) ----------

class MonthTab extends StatefulWidget {
  final _HomeShellState shell;
  const MonthTab({super.key, required this.shell});
  @override
  State<MonthTab> createState() => _MonthTabState();
}

class _MonthTabState extends State<MonthTab> {
  bool busy = false;
  _HomeShellState get sh => widget.shell;
  MoneyData get d => sh.d;

  Future<void> exportExcel() async {
    setState(() => busy = true);
    try {
      final bytes = buildMonthExcel(d, sh.month);
      if (bytes == null) throw 'Excel';
      final dir = await getAppFolder();
      final ex = Directory('${dir.path}/Excel');
      if (!await ex.exists()) await ex.create(recursive: true);
      final f = File('${ex.path}/A2R Money ${sh.month}.xlsx');
      await f.writeAsBytes(bytes, flush: true);
      if (!mounted) return;
      await showModalBottomSheet(
        context: context,
        builder: (c) => SafeArea(
          child: Padding(
            padding: const EdgeInsets.all(18),
            child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
              const Text('الشيت اتحفظ', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
              Text(f.path, style: const TextStyle(color: C.muted, fontSize: 12)),
              const SizedBox(height: 12),
              FilledButton.icon(
                  onPressed: () {
                    Navigator.pop(c);
                    openOrShareFile(context, f.path, kXlsxMime);
                  },
                  icon: const Icon(Icons.open_in_new_rounded),
                  label: const Text('افتح الشيت')),
              const SizedBox(height: 8),
              OutlinedButton.icon(
                  onPressed: () {
                    Navigator.pop(c);
                    Share.shareXFiles([XFile(f.path, mimeType: kXlsxMime)]);
                  },
                  icon: const Icon(Icons.share_rounded),
                  label: const Text('ابعته')),
            ]),
          ),
        ),
      );
    } catch (e) {
      if (mounted) toast(context, 'الشيت ماتعملش: $e', bad: true);
    } finally {
      if (mounted) setState(() => busy = false);
    }
  }

  Future<void> receiveMonth() async {
    setState(() => busy = true);
    await getAppFolder(requestPermissions: true);
    final files = await listMonthFiles(d.phoneId);
    if (!mounted) return;
    setState(() => busy = false);
    final pick = await showModalBottomSheet<FoundFile>(
      context: context,
      isScrollControlled: true,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 18, 16, 12),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            const Text('نسخ التليفون التاني اللي لقيناها', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            const Text('من واتساب و Download — نزّل الملف من الشات الأول', style: TextStyle(color: C.muted, fontSize: 12)),
            const SizedBox(height: 10),
            if (files.isEmpty)
              const Padding(
                padding: EdgeInsets.all(16),
                child: Text('مفيش ملفات. خلي التليفون التاني يدوس «ابعت نسختي» ويبعتها على واتساب، '
                    'وانت افتح الملف في الشات مرة (عشان يتنزل) وارجع هنا.',
                    style: TextStyle(color: C.muted)),
              ),
            Flexible(
              child: ListView(shrinkWrap: true, children: [
                for (final f in files)
                  ListTile(
                    leading: const Icon(Icons.description_rounded, color: C.gold),
                    title: Text(f.label, style: const TextStyle(fontWeight: FontWeight.w700)),
                    subtitle: Text('اتعمل ${arDate(f.when)} ${two(f.when.hour)}:${two(f.when.minute)}',
                        style: const TextStyle(color: C.muted, fontSize: 12)),
                    onTap: () => Navigator.pop(c, f),
                  ),
              ]),
            ),
          ]),
        ),
      ),
    );
    if (pick == null) return;
    final key = '${pick.json['phoneId']}|${pick.json['month']}';
    final old = pick.json['type'] == 'month' ? d.imported[key] : null;
    if (old != null && mounted) {
      final ok = await confirmDialog(context, 'الشهر ده متضاف قبل كده',
          'هيتبدل القديم (${old.txns.length} عملية) بالجديد — مفيش تكرار. تكمل؟', yes: 'بدّله');
      if (!ok) return;
    }
    final r = d.importCopyJson(pick.json);
    if (r.ok) {
      if (pick.json['type'] == 'month') sh.setMonth('${pick.json['month']}');
      await sh.persist();
    }
    if (mounted) toast(context, r.message, bad: !r.ok);
  }

  Widget _trRow(IconData ic, String label, int v, Color c) => Padding(
        padding: const EdgeInsets.symmetric(vertical: 3),
        child: Row(children: [
          Icon(ic, color: c, size: 20),
          const SizedBox(width: 8),
          Expanded(child: Text(label, style: const TextStyle(fontWeight: FontWeight.w700))),
          Text(egp(v), style: TextStyle(fontWeight: FontWeight.w900, color: c)),
        ]),
      );

  Widget _trWarn(String m) => Padding(
        padding: const EdgeInsets.only(top: 4, bottom: 2),
        child: Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
          const Icon(Icons.warning_amber_rounded, color: C.red, size: 18),
          const SizedBox(width: 6),
          Expanded(child: Text(m, style: const TextStyle(color: C.red, fontSize: 12))),
        ]),
      );

  @override
  Widget build(BuildContext context) {
    final mk = sh.month;
    final own = d.ownSummary(mk);
    final imports = d.importedFor(mk);
    final all = d.combinedSummary(mk);
    final tc = d.transferCheck(mk);
    final maxG = all.byGroup.values.fold<int>(1, max);
    return ListView(padding: const EdgeInsets.fromLTRB(16, 6, 16, 40), children: [
      MonthSwitcher(month: mk, onChanged: sh.setMonth),
      Container(
        padding: const EdgeInsets.all(18),
        decoration: BoxDecoration(
          borderRadius: BorderRadius.circular(24),
          gradient: const LinearGradient(colors: [Color(0xFF3B2A06), Color(0xFF0B3B38)], begin: Alignment.topRight, end: Alignment.bottomLeft),
          border: Border.all(color: C.gold.withValues(alpha: 0.6)),
        ),
        child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
          Text(imports.isEmpty ? 'إجمالي الصرف' : 'إجمالي الصرف (التليفونين)', style: const TextStyle(color: C.muted, fontWeight: FontWeight.w700)),
          FittedBox(
              fit: BoxFit.scaleDown,
              alignment: AlignmentDirectional.centerStart,
              child: Text(egp(all.totalSpend), style: const TextStyle(fontSize: 32, fontWeight: FontWeight.w900, color: C.gold))),
          const SizedBox(height: 8),
          Wrap(spacing: 8, runSpacing: 6, children: [
            Pill('مصروفات ${money(all.expenses)}', color: C.red),
            Pill('فواتير ${money(all.bills)}', color: C.gold),
            Pill('سحب ${money(all.withdraws)}', color: C.blue),
          ]),
        ]),
      ),
      const SectionTitle('حسب التليفون'),
      _phoneRow('${d.badgeName} (ده)', own, C.gold),
      for (final m in imports)
        _phoneRow('${m.phoneName} • استلمته ${arDateShort(m.exportedAt.length >= 10 ? m.exportedAt.substring(0, 10) : isoDate(DateTime.now()))}',
            d.importedSummary(m), C.emerald,
            onRemove: () async {
              final ok = await confirmDialog(context, 'تشيل ملف ${m.phoneName}؟', 'هيتشال من إجمالي ${monthLabel(mk)} بس',
                  yes: 'شيله', danger: true);
              if (!ok) return;
              d.imported.remove(m.key);
              await sh.persist();
            }),
      Padding(
        padding: const EdgeInsets.only(top: 6),
        child: Row(children: [
          Expanded(
            child: FilledButton.icon(
              style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 13)),
              onPressed: busy ? null : sh.syncNow,
              icon: const Icon(Icons.sync_rounded),
              label: const Text('تحديث', style: TextStyle(fontWeight: FontWeight.w800)),
            ),
          ),
          const SizedBox(width: 8),
          Expanded(
            child: OutlinedButton.icon(
              style: OutlinedButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 13)),
              onPressed: busy ? null : sh.sendCopy,
              icon: const Icon(Icons.ios_share_rounded),
              label: Text('ابعت نسختي لـ ${d.otherName}', maxLines: 1, overflow: TextOverflow.ellipsis),
            ),
          ),
        ]),
      ),
      Wrap(alignment: WrapAlignment.center, children: [
        TextButton.icon(
          onPressed: busy ? null : receiveMonth,
          icon: const Icon(Icons.download_rounded, size: 18),
          label: const Text('استلم ملف بإيدك'),
        ),
        TextButton.icon(
          onPressed: busy ? null : sh.receiveTransfer,
          icon: const Icon(Icons.install_mobile_rounded, size: 18, color: C.emerald),
          label: const Text('استلم تحويل'),
        ),
      ]),
      if (tc.sent > 0 || tc.received > 0 || tc.otherSent > 0 || tc.ownReceived > 0)
        Padding(
          padding: const EdgeInsets.only(top: 10),
          child: Container(
            padding: const EdgeInsets.all(12),
            decoration: BoxDecoration(
              color: C.card,
              borderRadius: BorderRadius.circular(18),
              border: Border.all(color: (tc.hasOther && (tc.sent != tc.received || tc.otherSent != tc.ownReceived) ? C.red : C.blue).withValues(alpha: 0.6)),
            ),
            child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
              if (tc.sent > 0 || tc.received > 0) ...[
                _trRow(Icons.send_to_mobile_rounded, 'انا حوّلت لـ ${d.otherName}', tc.sent, C.blue),
                if (tc.hasOther) _trRow(Icons.install_mobile_rounded, '${d.otherName} استلم', tc.received, C.emerald),
                if (tc.hasOther && tc.sent != tc.received)
                  _trWarn('فرق ${egp((tc.sent - tc.received).abs())} — غالباً في تحويل لسه ماتسجلش عند ${d.otherName} (أو اتسجل في شهر تاني).'),
              ],
              if (tc.otherSent > 0 || tc.ownReceived > 0) ...[
                if (tc.sent > 0 || tc.received > 0) const Divider(height: 14),
                if (tc.hasOther) _trRow(Icons.send_to_mobile_rounded, '${d.otherName} حوّل ليّا', tc.otherSent, C.blue),
                _trRow(Icons.install_mobile_rounded, 'انا استلمت', tc.ownReceived, C.emerald),
                if (tc.hasOther && tc.otherSent != tc.ownReceived)
                  _trWarn('فرق ${egp((tc.otherSent - tc.ownReceived).abs())} — دوس «تحديث» أو «استلم تحويل» عشان تسجّل اللي ناقص.'),
              ],
              const SizedBox(height: 6),
              const Text('التحويلات دي مش بتتحسب صرف — صرف كل تليفون بيتحسب مرة واحدة من نسخته.',
                  style: TextStyle(color: C.muted, fontSize: 11.5)),
            ]),
          ),
        ),
      if (all.byGroup.isNotEmpty) ...[
        const SectionTitle('حسب النوع'),
        for (final e in (all.byGroup.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
          _bar(e.key, e.value, maxG, e.key == kGroupHome ? C.gold : e.key == kGroupWork ? C.blue : e.key == kGroupPersonal ? C.emerald : C.violet),
      ],
      if (all.byCategory.isNotEmpty) ...[
        const SectionTitle('حسب البند'),
        Container(
          decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: C.line)),
          child: Column(children: [
            for (final e in (all.byCategory.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
              ListTile(dense: true, title: Text(e.key), trailing: Text(egp(e.value), style: const TextStyle(fontWeight: FontWeight.w800))),
          ]),
        ),
      ],
      if (all.byPerson.isNotEmpty) ...[
        const SectionTitle('الأسرة (صرف كل فرد)'),
        Container(
          decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: C.line)),
          child: Column(children: [
            for (final e in (all.byPerson.entries.toList()..sort((a, b) => b.value.compareTo(a.value))))
              ListTile(
                  dense: true,
                  leading: const Icon(Icons.person_rounded, color: C.emerald),
                  title: Text(e.key),
                  trailing: Text(egp(e.value), style: const TextStyle(fontWeight: FontWeight.w800))),
          ]),
        ),
      ],
      const SectionTitle('فلوس داخلة ورايحة'),
      Container(
        decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: C.line)),
        child: Column(children: [
          ListTile(dense: true, title: const Text('اتضاف رصيد'), trailing: Text(egp(all.added), style: const TextStyle(color: C.emerald, fontWeight: FontWeight.w800))),
          ListTile(dense: true, title: const Text('سلف اتدفعت'), trailing: Text(egp(all.advancesGiven), style: const TextStyle(color: C.violet, fontWeight: FontWeight.w800))),
          ListTile(dense: true, title: const Text('سلف رجعت'), trailing: Text(egp(all.advancesBack), style: const TextStyle(color: C.emerald, fontWeight: FontWeight.w800))),
          ListTile(dense: true, title: const Text('تحويلات بين الأماكن (مش صرف)'), trailing: Text(egp(all.transfers), style: const TextStyle(color: C.blue, fontWeight: FontWeight.w800))),
          if (own.toPhone > 0)
            ListTile(dense: true, title: Text('اتحول لـ ${d.otherName} (مش صرف)'), trailing: Text(egp(own.toPhone), style: const TextStyle(color: C.blue, fontWeight: FontWeight.w800))),
          if (own.fromPhone > 0)
            ListTile(dense: true, title: Text('اتاستلم من ${d.otherName} (مش صرف)'), trailing: Text(egp(own.fromPhone), style: const TextStyle(color: C.emerald, fontWeight: FontWeight.w800))),
        ]),
      ),
      const SizedBox(height: 16),
      FilledButton.tonalIcon(
        style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
        onPressed: busy ? null : exportExcel,
        icon: const Icon(Icons.table_view_rounded),
        label: const Text('طلّع شيت Excel للشهر'),
      ),
    ]);
  }

  Widget _phoneRow(String name, MonthSummary s, Color c, {VoidCallback? onRemove}) => Padding(
        padding: const EdgeInsets.only(bottom: 8),
        child: Container(
          padding: const EdgeInsets.fromLTRB(12, 10, 6, 10),
          decoration: BoxDecoration(color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: c.withValues(alpha: 0.5))),
          child: Row(children: [
            Icon(Icons.phone_android_rounded, color: c),
            const SizedBox(width: 10),
            Expanded(
                child: Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Text(name, style: const TextStyle(fontWeight: FontWeight.w800), maxLines: 1, overflow: TextOverflow.ellipsis),
              Text('${s.count} عملية', style: const TextStyle(color: C.muted, fontSize: 12)),
            ])),
            Text(egp(s.totalSpend), style: TextStyle(fontWeight: FontWeight.w900, color: c)),
            if (onRemove != null) IconButton(onPressed: onRemove, icon: const Icon(Icons.close_rounded, color: C.muted, size: 20)),
          ]),
        ),
      );

  Widget _bar(String l, int v, int maxV, Color c) => Padding(
        padding: const EdgeInsets.only(bottom: 10),
        child: Column(crossAxisAlignment: CrossAxisAlignment.stretch, children: [
          Row(children: [
            Text(l, style: const TextStyle(fontWeight: FontWeight.w700)),
            const Spacer(),
            Text(egp(v), style: TextStyle(fontWeight: FontWeight.w800, color: c)),
          ]),
          const SizedBox(height: 4),
          ClipRRect(
            borderRadius: BorderRadius.circular(8),
            child: LinearProgressIndicator(value: v / maxV, minHeight: 10, color: c, backgroundColor: C.bg2),
          ),
        ]),
      );
}

// ==================== ابدأ بأرصدتك ====================

/// النص اللي اتكتب ← قروش (يقبل سالب «-»، فاضي = صفر)
int? parseOpening(String raw) {
  final t = raw.trim();
  if (t.isEmpty) return 0;
  final neg = t.startsWith('-');
  final v = parseMoney(neg ? t.substring(1) : t);
  if (v == null) return null;
  return neg ? -v : v;
}

class StartBalancesScreen extends StatefulWidget {
  final MoneyData data;
  const StartBalancesScreen({super.key, required this.data});
  @override
  State<StartBalancesScreen> createState() => _StartBalancesScreenState();
}

class _StartBalancesScreenState extends State<StartBalancesScreen> {
  final ctl = <String, TextEditingController>{};
  final err = <String, String?>{};
  MoneyData get d => widget.data;

  @override
  void initState() {
    super.initState();
    for (final p in d.activePlaces) {
      ctl[p.id] = TextEditingController(text: p.opening == 0 ? '' : money(p.opening).replaceAll(',', ''));
    }
  }

  @override
  void dispose() {
    for (final c in ctl.values) {
      c.dispose();
    }
    super.dispose();
  }

  int get total => ctl.values.fold(0, (s, c) => s + (parseOpening(c.text) ?? 0));

  void save() {
    final vals = <String, int?>{};
    var bad = false;
    ctl.forEach((id, c) {
      final v = parseOpening(c.text);
      err[id] = v == null ? 'رقم مش مفهوم' : null;
      if (v == null) bad = true;
      vals[id] = v;
    });
    setState(() {});
    if (bad) return;
    d.setOpenings(vals);
    Navigator.pop(context, true);
  }

  @override
  Widget build(BuildContext context) {
    final places = d.activePlaces;
    return Scaffold(
      appBar: AppBar(title: const Text('ابدأ بأرصدتك')),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 8, 16, 120), children: [
        const Text('اكتب الفلوس اللي معاك دلوقتي في كل مكان. ده «رصيد البداية» — وبعد كده سجّل المصروف والسحب والإضافة عادي.',
            style: TextStyle(color: C.muted)),
        if (d.txns.isNotEmpty)
          const Padding(
            padding: EdgeInsets.only(top: 8),
            child: Text('خد بالك: عندك عمليات متسجلة — رصيد المكان = البداية + العمليات.',
                style: TextStyle(color: C.gold, fontSize: 12)),
          ),
        const SizedBox(height: 12),
        for (final p in places)
          Padding(
            padding: const EdgeInsets.only(bottom: 12),
            child: TextField(
              controller: ctl[p.id],
              keyboardType: const TextInputType.numberWithOptions(decimal: true, signed: true),
              inputFormatters: [FilteringTextInputFormatter.allow(RegExp(r'[0-9٠-٩۰-۹.,٫،\-]'))],
              onChanged: (_) => setState(() => err[p.id] = null),
              style: const TextStyle(fontSize: 20, fontWeight: FontWeight.w800),
              decoration: InputDecoration(
                labelText: p.name,
                prefixIcon: Icon(placeIcon(p.kind), color: placeColor(p.kind)),
                suffixText: 'جنيه',
                hintText: '0',
                errorText: err[p.id],
              ),
            ),
          ),
        Container(
          padding: const EdgeInsets.all(14),
          decoration: BoxDecoration(
              color: C.card, borderRadius: BorderRadius.circular(18), border: Border.all(color: C.gold.withValues(alpha: 0.5))),
          child: Row(children: [
            const Text('الإجمالي', style: TextStyle(fontWeight: FontWeight.w800)),
            const Spacer(),
            Text(egp(total), style: const TextStyle(fontSize: 20, fontWeight: FontWeight.w900, color: C.gold)),
          ]),
        ),
        const SizedBox(height: 8),
        const Text('عايز مكان تاني (فودافون كاش، بنك تاني…)؟ ضيفه من الإعدادات في جزء «الأماكن»، وهيظهر هنا.',
            style: TextStyle(color: C.muted, fontSize: 12)),
      ]),
      bottomNavigationBar: SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(16, 6, 16, 12),
          child: FilledButton.icon(
            style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 16)),
            onPressed: save,
            icon: const Icon(Icons.check_rounded),
            label: const Text('احفظ الأرصدة', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w800)),
          ),
        ),
      ),
    );
  }
}

// ==================== الإعدادات ====================

class SettingsScreen extends StatefulWidget {
  final _HomeShellState shell;
  const SettingsScreen({super.key, required this.shell});
  @override
  State<SettingsScreen> createState() => _SettingsScreenState();
}

class _SettingsScreenState extends State<SettingsScreen> {
  MoneyData get d => widget.shell.d;

  Future<void> save() async {
    await widget.shell.persist();
    if (mounted) setState(() {});
  }

  Future<void> editPlace({Place? p}) async {
    final name = TextEditingController(text: p?.name ?? '');
    final open = TextEditingController(text: p == null || p.opening == 0 ? '' : money(p.opening).replaceAll(',', ''));
    var kind = p?.kind ?? kPlaceOther;
    final r = await showDialog<bool>(
      context: context,
      builder: (c) => StatefulBuilder(
        builder: (c, set) => AlertDialog(
          title: Text(p == null ? 'مكان جديد' : 'تعديل ${p.name}', style: const TextStyle(fontWeight: FontWeight.w800)),
          content: SingleChildScrollView(
            child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
              TextField(controller: name, decoration: const InputDecoration(labelText: 'الاسم (مثلاً: فودافون كاش، البنك الأهلي)')),
              const SizedBox(height: 10),
              Wrap(spacing: 6, runSpacing: 6, children: [
                for (final k in const [
                  [kPlaceCash, 'كاش'],
                  [kPlaceBank, 'بنك'],
                  [kPlaceWallet, 'محفظة'],
                  [kPlaceOther, 'تاني'],
                ])
                  ChoiceChip(
                    avatar: Icon(placeIcon(k[0]), size: 16, color: placeColor(k[0])),
                    label: Text(k[1]),
                    selected: kind == k[0],
                    showCheckmark: false,
                    onSelected: (_) => set(() => kind = k[0]),
                  ),
              ]),
              const SizedBox(height: 10),
              TextField(
                controller: open,
                keyboardType: const TextInputType.numberWithOptions(decimal: true, signed: true),
                decoration: const InputDecoration(labelText: 'رصيد البداية', suffixText: 'جنيه'),
              ),
              const SizedBox(height: 6),
              const Text('رصيد البداية = الفلوس اللي كانت في المكان قبل ما تبدأ تسجّل. بعد كده استخدم «إضافة رصيد» و«سحب».',
                  style: TextStyle(color: C.muted, fontSize: 12)),
            ]),
          ),
          actions: [
            TextButton(onPressed: () => Navigator.pop(c, false), child: const Text('إلغاء')),
            FilledButton(onPressed: () => Navigator.pop(c, true), child: const Text('حفظ')),
          ],
        ),
      ),
    );
    if (r != true || name.text.trim().isEmpty) return;
    final raw = open.text.trim();
    final neg = raw.startsWith('-');
    final v = parseMoney(neg ? raw.substring(1) : raw) ?? 0;
    if (p == null) {
      d.places.add(Place(id: newId('pl'), name: name.text.trim(), kind: kind, opening: neg ? -v : v));
    } else {
      p.name = name.text.trim();
      p.kind = kind;
      p.opening = neg ? -v : v;
    }
    await save();
  }

  Future<void> removePlace(Place p) async {
    final used = d.txns.any((t) => t.placeId == p.id || t.toPlaceId == p.id);
    if (used) {
      final ok = await confirmDialog(context, 'المكان ده عليه عمليات',
          'مش هيتمسح عشان الحساب مايتلخبطش — هيتخبى من الاختيارات (رصيده الحالي ${egp(d.balanceOf(p.id))} هيفضل في الإجمالي). تخبيه؟',
          yes: 'خبيه');
      if (!ok) return;
      p.archived = true;
    } else {
      final ok = await confirmDialog(context, 'تمسح «${p.name}»؟', 'مفيش عليه عمليات', yes: 'امسح', danger: true);
      if (!ok) return;
      d.places.remove(p);
    }
    await save();
  }

  Future<void> editPerson({Person? p}) async {
    final name = TextEditingController(text: p?.name ?? '');
    var fam = p?.family ?? true;
    var child = p?.child ?? false;
    final ok = await showDialog<bool>(
      context: context,
      builder: (c) => StatefulBuilder(
        builder: (c, set) => AlertDialog(
          title: Text(p == null ? 'شخص جديد' : 'تعديل', style: const TextStyle(fontWeight: FontWeight.w800)),
          content: Column(mainAxisSize: MainAxisSize.min, children: [
            TextField(controller: name, autofocus: true, decoration: const InputDecoration(labelText: 'الاسم')),
            SwitchListTile(
                contentPadding: EdgeInsets.zero,
                title: const Text('من الأسرة'),
                subtitle: const Text('بيظهر في «لمين؟» وفي تقرير الأسرة', style: TextStyle(fontSize: 12, color: C.muted)),
                value: fam,
                onChanged: (v) => set(() => fam = v)),
            SwitchListTile(contentPadding: EdgeInsets.zero, title: const Text('طفل'), value: child, onChanged: (v) => set(() => child = v)),
          ]),
          actions: [
            TextButton(onPressed: () => Navigator.pop(c, false), child: const Text('إلغاء')),
            FilledButton(onPressed: () => Navigator.pop(c, true), child: const Text('حفظ')),
          ],
        ),
      ),
    );
    final nm = cleanLabel(name.text);
    if (ok != true || nm.isEmpty) return;
    final dup = d.people.any((x) => x != p && sameLabel(x.name, nm));
    if (dup) {
      if (mounted) toast(context, '«$nm» موجود قبل كده في القايمة', bad: true);
      return;
    }
    if (p == null) {
      d.people.add(Person(id: newId('pe'), name: nm, family: fam, child: child));
    } else {
      p.name = nm;
      p.family = fam;
      p.child = child;
    }
    await save();
  }

  Future<void> removePerson(Person p) async {
    if (d.txns.any((t) => t.personId == p.id)) {
      toast(context, 'مينفعش يتمسح — عليه عمليات. تقدر تغيّر اسمه بس', bad: true);
      return;
    }
    final ok = await confirmDialog(context, 'تمسح ${p.name}؟', '', yes: 'امسح', danger: true);
    if (!ok) return;
    d.people.remove(p);
    await save();
  }

  Future<void> manageCategories(String g) async {
    await Navigator.push(
        context,
        MaterialPageRoute(
            builder: (_) => ChoicesScreen(
                  title: 'بنود $g',
                  list: d.categories.putIfAbsent(g, () => []),
                  addTitle: 'بند جديد في «$g»',
                  deleteNote: 'العمليات القديمة هتفضل باسم البند ده.',
                  onRename: (a, b) {
                    for (final t in d.txns) {
                      if (t.group == g && t.category == a) t.category = b;
                    }
                  },
                )));
    await save();
  }

  Future<void> manageList(String title, List<String> list, String addTitle) async {
    await Navigator.push(
        context, MaterialPageRoute(builder: (_) => ChoicesScreen(title: title, list: list, addTitle: addTitle)));
    await save();
  }

  Future<void> backupNow() async {
    await getAppFolder(requestPermissions: true);
    final f = await writeAutoBackup(d);
    try {
      final dir = await getAppFolder(requestPermissions: false);
      final m = File('${dir.path}/$kManualBackupName');
      await m.writeAsString(jsonEncode(backupPayload(d)), flush: true);
      if (!mounted) return;
      final share = await confirmDialog(context, 'اتعملت نسخة', 'اتحفظت في ${m.path}\nتبعتها لنفسك على واتساب كمان؟',
          yes: 'ابعتها', no: 'خلاص');
      if (share) await Share.shareXFiles([XFile(m.path, mimeType: 'application/json')]);
    } catch (e) {
      if (mounted) toast(context, f == null ? 'النسخة ماتعملتش — اسمح بصلاحية الملفات' : 'اتعملت نسخة تلقائية', bad: f == null);
    }
  }

  Future<void> restore() async {
    await getAppFolder(requestPermissions: true);
    final list = await listBackups();
    if (!mounted) return;
    if (list.isEmpty) {
      toast(context, 'مفيش نسخ لقيناها', bad: true);
      return;
    }
    final pick = await showModalBottomSheet<FoundFile>(
      context: context,
      isScrollControlled: true,
      builder: (c) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.all(16),
          child: Column(mainAxisSize: MainAxisSize.min, crossAxisAlignment: CrossAxisAlignment.stretch, children: [
            const Text('اختار النسخة', style: TextStyle(fontSize: 18, fontWeight: FontWeight.w800)),
            Flexible(
              child: ListView(shrinkWrap: true, children: [
                for (final b in list)
                  ListTile(leading: const Icon(Icons.backup_rounded, color: C.emerald), title: Text(b.label), onTap: () => Navigator.pop(c, b)),
              ]),
            ),
          ]),
        ),
      ),
    );
    if (pick == null || !mounted) return;
    final ok = await confirmDialog(context, 'ترجّع النسخة دي؟',
        '${pick.label}\n\nالبيانات الحالية هتتحفظ الأول في نسخة «قبل الاسترجاع».',
        yes: 'رجّعها', danger: true);
    if (!ok) return;
    await writeSafetyBackup(d);
    final nd = MoneyData.fromJson(pick.json);
    if (nd.role.isEmpty) nd.role = d.role;
    if (!mounted) return;
    Navigator.pop(context);
    await widget.shell.widget.onReplace(nd);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: const Text('الإعدادات')),
      body: ListView(padding: const EdgeInsets.fromLTRB(16, 4, 16, 40), children: [
        SectionTitle('الأماكن (الأرصدة)',
            trailing: TextButton.icon(onPressed: () => editPlace(), icon: const Icon(Icons.add_rounded), label: const Text('مكان'))),
        for (final p in d.places)
          ListTile(
            leading: Icon(placeIcon(p.kind), color: p.archived ? C.muted : placeColor(p.kind)),
            title: Text(p.name + (p.archived ? ' (مخفي)' : ''), style: const TextStyle(fontWeight: FontWeight.w700)),
            subtitle: Text('الرصيد ${egp(d.balanceOf(p.id))} • البداية ${egp(p.opening)}', style: const TextStyle(color: C.muted, fontSize: 12)),
            onTap: () => editPlace(p: p),
            trailing: p.archived
                ? TextButton(
                    onPressed: () async {
                      p.archived = false;
                      await save();
                    },
                    child: const Text('رجّعه'))
                : IconButton(icon: const Icon(Icons.delete_outline_rounded), onPressed: () => removePlace(p)),
          ),
        OutlinedButton.icon(
          onPressed: () async {
            final ok = await Navigator.push<bool>(context, MaterialPageRoute(builder: (_) => StartBalancesScreen(data: d)));
            if (ok == true) await save();
          },
          icon: const Icon(Icons.edit_note_rounded),
          label: const Text('اكتب رصيد البداية لكل الأماكن مرة واحدة'),
        ),
        SectionTitle('الأسرة والناس',
            trailing: TextButton.icon(onPressed: () => editPerson(), icon: const Icon(Icons.add_rounded), label: const Text('شخص'))),
        if (d.people.isEmpty) const Text('ضيف الأطفال وأفراد الأسرة عشان تعرف صرف كل واحد', style: TextStyle(color: C.muted)),
        for (final p in d.people)
          ListTile(
            leading: Icon(p.child ? Icons.child_care_rounded : (p.family ? Icons.family_restroom_rounded : Icons.person_outline_rounded),
                color: p.child ? C.gold : C.emerald),
            title: Text(p.name),
            subtitle: Text(p.family ? (p.child ? 'طفل' : 'من الأسرة') : 'مش من الأسرة (سلف)', style: const TextStyle(color: C.muted, fontSize: 12)),
            onTap: () => editPerson(p: p),
            trailing: IconButton(icon: const Icon(Icons.delete_outline_rounded), onPressed: () => removePerson(p)),
          ),
        const SectionTitle('بنود المصروفات'),
        for (final g in kGroups)
          ListTile(
            leading: const Icon(Icons.category_rounded, color: C.gold),
            title: Text('بنود $g'),
            subtitle: Text((d.categories[g] ?? []).join('، '), maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: C.muted, fontSize: 12)),
            onTap: () => manageCategories(g),
          ),
        const SectionTitle('الاختيارات المحفوظة'),
        const Text('أي حاجة بتكتبها جديدة وانت بتسجل بتتحفظ هنا لوحدها — تقدر تعدلها أو تمسحها.',
            style: TextStyle(color: C.muted, fontSize: 12)),
        ListTile(
          leading: const Icon(Icons.south_west_rounded, color: C.emerald),
          title: const Text('الفلوس جاية منين (إضافة رصيد)'),
          subtitle: Text(d.sources.isEmpty ? 'لسه مفيش' : d.sources.join('، '),
              maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: C.muted, fontSize: 12)),
          onTap: () => manageList('الفلوس جاية منين', d.sources, 'مصدر جديد'),
        ),
        ListTile(
          leading: const Icon(Icons.outbox_rounded, color: C.red),
          title: const Text('أسباب السحب'),
          subtitle: Text(d.reasons.isEmpty ? 'لسه مفيش' : d.reasons.join('، '),
              maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: C.muted, fontSize: 12)),
          onTap: () => manageList('أسباب السحب', d.reasons, 'سبب جديد'),
        ),
        ListTile(
          leading: const Icon(Icons.receipt_long_rounded, color: C.gold),
          title: const Text('أسماء الفواتير'),
          subtitle: Text(d.billNames.isEmpty ? 'لسه مفيش' : d.billNames.join('، '),
              maxLines: 1, overflow: TextOverflow.ellipsis, style: const TextStyle(color: C.muted, fontSize: 12)),
          onTap: () => manageList('أسماء الفواتير', d.billNames, 'اسم فاتورة جديد'),
        ),
        const SectionTitle('القوايم بين التليفونين'),
        const Text('ابعت الأسرة والبنود والفواتير والأماكن لتليفون تاني بدل ما تكتبهم من جديد. '
            'من غير أي فلوس: لا أرصدة ولا عمليات ولا سلف.',
            style: TextStyle(color: C.muted, fontSize: 12)),
        const SizedBox(height: 8),
        FilledButton.tonalIcon(
          onPressed: () => sendListsFile(context, d),
          icon: const Icon(Icons.ios_share_rounded),
          label: const Text('ابعت القوايم لتليفون تاني'),
        ),
        const SizedBox(height: 8),
        OutlinedButton.icon(
          onPressed: () async {
            final changed = await receiveListsFlow(context, d);
            if (changed) await save();
          },
          icon: const Icon(Icons.playlist_add_rounded),
          label: const Text('استلم القوايم من تليفون تاني'),
        ),
        const SectionTitle('التليفون ده'),
        ListTile(
          leading: const Icon(Icons.badge_rounded, color: C.emerald),
          title: const Text('اسم التليفون'),
          subtitle: Text(d.phoneName, style: const TextStyle(color: C.muted)),
          onTap: () async {
            final n = await askText(context, 'اسم التليفون', initial: d.phoneName);
            if (n == null) return;
            d.phoneName = n;
            await save();
          },
        ),
        ListTile(
          leading: const Icon(Icons.folder_special_rounded, color: C.gold),
          title: const Text('فولدر الاستلام (للتحديث)'),
          subtitle: Text(d.syncFolder.isEmpty ? 'لسه ماتحددش — هيتسأل أول مرة تدوس «تحديث»' : d.syncFolder,
              style: const TextStyle(color: C.muted, fontSize: 12)),
          onTap: () async {
            await widget.shell.chooseSyncFolder();
            if (mounted) setState(() {});
          },
        ),
        const SectionTitle('النسخ الاحتياطي'),
        const Text('نسخة تلقائية كل يوم في «A2R Money/Backups» (آخر ١٤ يوم).', style: TextStyle(color: C.muted, fontSize: 12)),
        const SizedBox(height: 8),
        FilledButton.tonalIcon(onPressed: backupNow, icon: const Icon(Icons.save_rounded), label: const Text('اعمل نسخة دلوقتي')),
        const SizedBox(height: 8),
        OutlinedButton.icon(onPressed: restore, icon: const Icon(Icons.restore_rounded), label: const Text('رجّع نسخة')),
        const SizedBox(height: 24),
        const Center(child: BrandHeader(small: true)),
        const SizedBox(height: 4),
        const Center(child: Text('الإصدار $kAppVersion', style: TextStyle(color: C.muted, fontSize: 12))),
      ]),
    );
  }
}

/// قايمة اختيارات (بنود / مصادر الفلوس / أسباب السحب / أسماء الفواتير): إضافة، تعديل، مسح، ترتيب — من غير تكرار
class ChoicesScreen extends StatefulWidget {
  final String title;
  final List<String> list;
  final String addTitle;
  final String deleteNote;
  final void Function(String oldName, String newName)? onRename;
  const ChoicesScreen({
    super.key,
    required this.title,
    required this.list,
    this.addTitle = 'اختيار جديد',
    this.deleteNote = 'العمليات القديمة هتفضل زي ما هي.',
    this.onRename,
  });
  @override
  State<ChoicesScreen> createState() => _ChoicesScreenState();
}

class _ChoicesScreenState extends State<ChoicesScreen> {
  List<String> get list => widget.list;

  bool _taken(String n, {String? except}) =>
      list.any((x) => x != except && sameLabel(x, n));

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(title: Text(widget.title)),
      floatingActionButton: FloatingActionButton(
        onPressed: () async {
          final n = await askText(context, widget.addTitle);
          if (n == null) return;
          if (_taken(n)) {
            if (context.mounted) toast(context, '«${cleanLabel(n)}» موجود قبل كده');
            return;
          }
          setState(() => addChoice(list, n));
        },
        child: const Icon(Icons.add_rounded),
      ),
      body: list.isEmpty
          ? const Center(
              child: Padding(
              padding: EdgeInsets.all(24),
              child: Text('لسه مفيش — أول ما تكتب حاجة جديدة في التسجيل هتتحفظ هنا لوحدها',
                  textAlign: TextAlign.center, style: TextStyle(color: C.muted)),
            ))
          : ReorderableListView(
              padding: const EdgeInsets.fromLTRB(16, 6, 16, 90),
              onReorderItem: (a, b) => setState(() => list.insert(b, list.removeAt(a))),
              children: [
                for (final c in list)
                  ListTile(
                    key: ValueKey(c),
                    title: Text(c),
                    leading: IconButton(
                      tooltip: 'تعديل',
                      icon: const Icon(Icons.edit_rounded, size: 20),
                      onPressed: () async {
                        final n0 = await askText(context, 'تغيير الاسم', initial: c);
                        if (n0 == null) return;
                        final n = cleanLabel(n0);
                        if (n.isEmpty || n == c) return;
                        if (_taken(n, except: c)) {
                          if (context.mounted) toast(context, '«$n» موجود قبل كده');
                          return;
                        }
                        setState(() {
                          list[list.indexOf(c)] = n;
                          widget.onRename?.call(c, n);
                        });
                      },
                    ),
                    trailing: IconButton(
                      tooltip: 'مسح',
                      icon: const Icon(Icons.delete_outline_rounded),
                      onPressed: () async {
                        final ok = await confirmDialog(context, 'تمسح «$c»؟', widget.deleteNote, yes: 'امسح', danger: true);
                        if (ok) setState(() => list.remove(c));
                      },
                    ),
                  ),
              ],
            ),
    );
  }
}
