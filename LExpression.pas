unit LExpression;

interface

uses
  Classes, RegularExpressions, System.Generics.Collections, SysUtils,
  Variants, DateUtils, Math;

const
  ERR_UNKNOWN_VARIABLE = 'Неизвестная переменная: %s';
  ERR_UNKNOWN_FUNCTION = 'Неизвестная функция: %s';
  ERR_UNDEFINED_OPERATION = 'Операция "%s %s %s" не определена';
  ERR_EXPRESSION_NOT_FOUND = 'Выражение %s не найдено';
  ERR_CIRCULAR_REFERENCE = 'Ссылки выражений друг на друга';

type
  //Типы элементов выражений
  TItemType = (itInteger, itFloat, itDate, itString, itExpression, itVariable, itNull, itList, itFunction);
  TExpressionType = (etExpression, etFunction);

  TGetPriorityFunc = function(aOperation: String): Byte;

  TLExpression = class
  private
    FName: String;
    FExpression: String;
    FValue: Variant;
    FItemType: TItemType;
    FOperation: String;
    FPriority: Byte;

    FSiblings: TObjectList<TLExpression>;
    //Разбор выражения
    procedure parseExpression(aExprStr: String; aExprType: TExpressionType);
    //Разбор функции
    procedure parseFunction(aFuncStr: String);
    //Автоопределение типа выражения
    procedure defineType;
    function isFunction(aStr: String): boolean;
    procedure evaluate(aUsedVars: String);
    //Поиск связанных переменных
    function getSiblingByName(aName: String): TLExpression;
    //Функция определяющая следующую по приоритету операцию
    function getNextOperationItem: Integer;
    //Сохранение операции и её приоритета
    procedure setOperation(aOper: String);

    procedure processOperation(aItem1, aItem2: TLExpression);
 protected
    FChildren: TObjectList<TLExpression>;

    constructor Create(aExpression: String; aSiblings: TObjectList<TLExpression>; aOperation: String = ''; aName: String = ''); overload;
    function getValue: Variant;

    procedure setNullValueByType(aItem: TLExpression; aType: TItemType); virtual;
    //Определяет приоритет операций (0 - не операция,)
    function getOperationPriority(aOperation: String): Byte; virtual;

    function isFunctionDelimiter(aInput: String): Byte; virtual;
    //При наследовании обязательно должен переопределяться соответствующим контруктором дочернего класса
    function getInstance(aExpression: String; aSiblings: TObjectList<TLExpression>; aOperation: String = ''; aName: String = ''): TLExpression; virtual;

    //Функция вычисления операций
    function performOperation(aItem1, aItem2: TLExpression): boolean; virtual;

    //Вычисляет результат функции и записывает в FValue
    function performFunction: boolean; virtual;
  public
    //Конструктор для списка выражений
    constructor Create; overload;
    destructor Destroy; override;
    //Добавить выражение в список
    procedure AddExpression(aName, aExpression: String);
    //Очистка обеъкта выражения
    procedure Clear;

    //Получить результат выражения из списка
    function CalcValue(aExpressionName: String): String;
    //Получить данного выражения
    property Value: Variant read getValue write FValue;
    property Operation: String read FOperation write setOperation;
    property Expression: String read FExpression;
    property ItemType: TItemType read FItemType write FItemType;
    property Siblings: TObjectList<TLExpression> read FSiblings;
    property Name: String read FName;
  end;

var
  MaxPriority: Word;

implementation

{ TLExpression }

constructor TLExpression.Create;
begin
  inherited Create;

  FChildren := TObjectList<TLExpression>.Create;
  FSiblings := TObjectList<TLExpression>.Create;
  FItemType := itNull;
  Operation := '';
end;

constructor TLExpression.Create(aExpression: String; aSiblings: TObjectList<TLExpression>; aOperation: String = ''; aName: String = '');
var
  priorityFunc: TGetPriorityFunc;
begin
  Create;

  FName := aName;
  Operation := aOperation;
  FSiblings := aSiblings;

  //Удаление скобок по краям
  FExpression := trim(aExpression);
  if (length(FExpression) > 1) and (FExpression[1] = '(') and (FExpression[length(FExpression)] = ')') then
  begin
    Delete(FExpression, length(FExpression), 1);
    Delete(FExpression, 1, 1);
  end;

  defineType;

  if FItemType = itExpression then
  begin
    parseExpression(FExpression, etExpression);
  end;
  if FItemType = itFunction then
  begin
    parseFunction(FExpression);
  end;
end;

destructor TLExpression.Destroy;
begin
  FChildren.Free;

  //Удаляем список выражений
  if (FItemType = itList) then
    FSiblings.Free;

  inherited;
end;

procedure TLExpression.AddExpression(aName, aExpression: String);
var
  aExpr: TLExpression;
begin
  FItemType := itList;

  aExpr := GetInstance(aExpression, FSiblings, '', aName);
  //Создаем список выражений (один раз)
  if not Assigned(FSiblings) then
    FSiblings := TObjectList<TLExpression>.Create;
  FSiblings.Add(aExpr);
end;

function TLExpression.CalcValue(aExpressionName: String): String;
var
  aExprObj: TLExpression;
begin
  aExprObj := getSiblingByName(aExpressionName);
  if not Assigned(aExprObj) then
  begin
    result := format(ERR_EXPRESSION_NOT_FOUND, [aExpressionName]);
  end
  else
  begin
    result := aExprObj.getValue;
  end;
end;


procedure TLExpression.Clear;
begin
  FSiblings.Clear;
  FChildren.Clear;
end;

procedure TLExpression.defineType;
var
  i: Integer;
  inQuotes: Boolean;
  intVal: Integer;
  dateVal: TDateTime;
  floatVal: Double;
  rx: TRegEx;
begin
  if FExpression = '' then
  begin
    FItemType := itNull;
    exit;
  end;

  if isFunction(FExpression) then
  begin
    FItemType := itFunction;
    exit;
  end;

  //Если найдены знаки операций не внутри кавычек
  inQuotes := false;
  for i := 1 to length(FExpression) do
  begin
    if FExpression[i] = Chr(39) then
      inQuotes := not inQuotes;

    if (not inQuotes) and (getOperationPriority(FExpression[i]) > 0) then
    begin
      FItemType := itExpression;
      exit;
    end;
  end;

  //Если краям #, считаем переменной
  if (length(FExpression) > 2) and (FExpression[1] = '#') and (FExpression[length(FExpression)] = '#') then
  begin
    FItemType := itVariable;
    //Убираем #
    FExpression := copy(FExpression, 2, length(FExpression) - 2);
    exit;
  end;

  if TryStrToInt(FExpression, intVal) then
  begin
    FItemType := itInteger;
    FValue := intVal;
  end
  else if TryStrToFloat(FExpression, floatVal) then
  begin
    FItemType := itFloat;
    FValue := floatVal;
  end
  else if TryStrToDateTime(FExpression, dateVal) then
  begin
    FItemType := itDate;
    FValue := dateVal;
  end
  else
  begin
    FItemType := itString;

        //Удаление кавычек по краям
    if (length(FExpression) > 1) and (FExpression[1] = Chr(39)) and (FExpression[length(FExpression)] = Chr(39)) then
    begin
      Delete(FExpression, length(FExpression), 1);
      Delete(FExpression, 1, 1);
    end;
    FValue := FExpression;
  end;
end;

procedure TLExpression.evaluate(aUsedVars: String);
var
  i: Integer;
  aExprObj: TLExpression;
  index: Integer;
begin
  case FItemType of
  itVariable:
    begin
      if pos(FName + ',', aUsedVars) > 0 then
        raise Exception.Create(ERR_CIRCULAR_REFERENCE);

      if FName <> '' then
        aUsedVars := aUsedVars + FName + ',';

      aExprObj := getSiblingByName(FExpression);

      //Если переменная не найдена
      if not Assigned(aExprObj) then
      begin
        raise Exception.Create(format(ERR_UNKNOWN_VARIABLE,[FExpression]));
        exit;
      end;

      //Попытка вычислить значение переменной
      aExprObj.evaluate(aUsedVars);
      FValue := aExprObj.FValue;
      FItemType := aExprObj.FItemType;
      exit;
    end;
  itFunction:
    begin
      if not performFunction then
      begin
        raise Exception.Create(format(ERR_UNKNOWN_FUNCTION,[FExpression]));
      end;
    end;
  itExpression:
    begin
      //Вычиисление отдельных элементов
      for i := 0 to FChildren.Count - 1 do
        FChildren[i].evaluate(aUsedVars);

        //Вычисление операций
      while FChildren.Count > 1 do
      begin
        index := getNextOperationItem;
        processOperation(FChildren[index], FChildren[index+1]);
        FChildren.Delete(index);
      end;
      //Устанавливаем тип результата
      FItemType := FChildren[0].FItemType;
      //Сохраняем результат вычисления (чтобы не вычислять заново)
      FValue := FChildren[0].FValue;
    end;
  end;

  //Удаление дочерних выражений
  FChildren.Clear;
end;

function TLExpression.getInstance(aExpression: String; aSiblings: TObjectList<TLExpression>; aOperation: String = ''; aName: String = ''): TLExpression;
begin
  result := TLExpression.Create(aExpression, aSiblings, aOperation, aName);
end;

function TLExpression.getNextOperationItem: Integer;
var
  i: Integer;
  aPriority: Word;
begin
  result := -1;

  for aPriority := MaxPriority downto 1 do
  begin
    for i := 0 to FChildren.Count - 1 do
    begin
      if (FChildren[i].FPriority = aPriority) then
      begin
        result := i;
        exit;
      end;
    end;
  end;
end;

function TLExpression.getOperationPriority(aOperation: String): Byte;
begin
  //Отделяет аргументы функции
  if (aOperation = ';') then
    result := 3
  else if (aOperation = '*') or (aOperation = '/') then
    result := 2
  else if (aOperation = '+') or (aOperation = '-') then
    result := 1
  else
    result := 0;
end;

function TLExpression.isFunctionDelimiter(aInput: String): Byte;
begin
  if aInput = ';' then
    result := 1
  else
    result := 0;
end;

function TLExpression.getSiblingByName(aName: String): TLExpression;
var
  i: Integer;
begin
  result := nil;

  //Если список других переменных не указан
  if not Assigned(FSiblings) then
    exit;

  for i := 0 to FSiblings.Count - 1 do
  begin
    if (FSiblings[i].FName = aName) then
    begin
      result := FSiblings[i];
      exit;
    end;
  end;
end;

function TLExpression.getValue: Variant;
begin
  evaluate('');

  result := FValue;
end;

function TLExpression.isFunction(aStr: String): boolean;
var
  i: Integer;
  level: byte;
  rx: TRegEx;
begin
  if not (rx.IsMatch(aStr, '^[a-zA-Z]+\(')) then
  begin
    result := false;
    exit;
  end;

  level := 1;

  for i := pos('(', aStr) + 1 to length(aStr) do
  begin
    if aStr[i] = '(' then
      inc(level);
    if aStr[i] = ')' then
      dec(level);

    if (level = 0) then
    begin
      if (i < length(aStr)) then
        result := false
      else
        result := true;
      exit;
    end;
  end;
end;


procedure TLExpression.parseExpression(aExprStr: String; aExprType: TExpressionType);
var
  i: Integer;
  curItem: String;
  inQuotes: Boolean;

  bracketLevel: Integer;
  newItem: TLExpression;

  isOper: boolean;
begin
  FChildren.Clear;
  bracketLevel := 0;

  for i := 1 to length(aExprStr) do
  begin
    if (aExprStr[i] = Chr(39)) then
      inQuotes := not inQuotes;

    if (aExprStr[i] = '(') then
    begin
      inc(bracketLevel);
    end;

    if (aExprStr[i] = ')') then
    begin
      dec(bracketLevel);
    end;

    if aExprType = etFunction then
      isOper := (bracketLevel <= 0) and (not inQuotes) and (isFunctionDelimiter(aExprStr[i]) > 0)
    else
      isOper := (bracketLevel <= 0) and (not inQuotes) and (getOperationPriority(aExprStr[i]) > 0);
    if isOper then
    begin
      newItem := GetInstance(curItem, FSiblings, aExprStr[i]);
      FChildren.Add(newItem);

      curItem := '';
      continue;
    end;

    curItem := curItem + aExprStr[i];
  end;
  newItem := GetInstance(curItem, FSiblings);
  FChildren.Add(newItem);
end;

procedure TLExpression.parseFunction(aFuncStr: String);
var
  rx: TRegEx;
  match: TMatch;
begin
  match := rx.Match(aFuncStr, '([a-zA-Z]+)\((.*)\)');
  FExpression := match.Groups.Item[1].Value;

  parseExpression(match.Groups.Item[2].Value, etFunction);
end;

function TLExpression.performFunction: boolean;
var
  i: Integer;
  intVal1, intVal2: Integer;
  floatVal1, floatVal2: Double;
begin
  result := true;

  if SameText(FExpression, 'pi') then
  begin
    FValue := 3.14;
    FItemType := itFloat;
  end
  else if SameText(FExpression, 'randomStr') then
  begin
    FValue := FChildren[Random(FChildren.Count)].getValue;
    FItemType := itString;
  end
  else if SameText(FExpression, 'randomInt') then
  begin
    if FChildren.Count = 1 then
    begin
      intVal1 := FChildren[0].getValue;
      FValue := Random(intVal1);
    end
    else if FChildren.Count = 2 then
    begin
      intVal1 := FChildren[0].getValue;
      intVal2 := FChildren[1].getValue;
      FValue := RandomRange(intVal1, intVal2);
    end
    else
      raise Exception.Create('Неверное кол-во аргументов в функции randomInt');

    FItemType := itInteger;
  end
  else if SameText(FExpression, 'randomFloat') then
  begin
    if FChildren.Count = 1 then
    begin
      floatVal1 := FChildren[0].getValue;
      FValue := floatVal1 * random;
    end
    else if FChildren.Count = 2 then
    begin
      floatVal1 := FChildren[0].getValue;
      floatVal2 := FChildren[1].getValue;
      FValue := floatVal1 + (floatVal2 - floatVal1)*random;
    end
    else
      raise Exception.Create('Неверное кол-во аргументов в функции randomFloat');

    FItemType := itFloat;
  end
  else if SameText(FExpression, 'today') then
  begin
    FValue := Today;
    FItemType := itDate;
  end
  else if Sametext(FExpression, 'now') then
  begin
    FValue := Now;
    FItemType := itDate;
  end
  else
  begin
    result := false;
  end;
end;

function TLExpression.performOperation(aItem1, aItem2: TLExpression): boolean;
begin
  result := true;

  //Операции для чисел и дат
  if (aItem1.FItemType in [itInteger, itFloat, itDate]) and (aItem2.FItemType in [itInteger, itFloat, itDate]) then
  begin
    if aItem1.FOperation = '+' then
      aItem2.FValue := aItem1.FValue + aItem2.FValue
    else if aItem1.FOperation = '-' then
      aItem2.FValue := aItem1.FValue - aItem2.FValue
    else if aItem1.FOperation = '*' then
      aItem2.FValue := aItem1.FValue * aItem2.FValue
    else if aItem1.FOperation = '/' then
      aItem2.FValue := aItem1.FValue / aItem2.FValue
    else
      aItem2.FValue := 0;
  end
  //Если хотя один операнд = строка, выполняем операции для строк
  else if (aItem1.FItemType = itString) or (aItem2.FItemType = itString) then
  begin
    if aItem1.FOperation = '+' then
      aItem2.FValue := VarToStr(aItem1.FValue) + VarToStr(aItem2.Value)
    else if aItem1.FOperation = '-' then
      aItem2.FValue := StringReplace(VarToStr(aItem1.FValue), VarToStr(aItem2.FValue), '', [rfReplaceAll])
    else
      aItem2.FValue := '';
  end
  else
  begin
    result := false;
    exit;
  end;

  //Первый операнд всегда обнуляется (для порядка)
  aItem1.FValue := '';
  aItem1.FItemType := itNull;
end;

procedure TLExpression.processOperation(aItem1, aItem2: TLExpression);
var
  i: Integer;
begin
   //Если один из операндов пустой
  if (aItem1.FItemType = itNull) then
  begin
    setNullValueByType(aItem1, aItem2.FItemType);
  end;

  if (aItem2.FItemType = itNull) then
  begin
    setNullValueByType(aItem2, aItem1.FItemType);
  end;

  if (not performOperation(aItem1, aItem2)) then
  begin
    raise Exception.Create(format(ERR_UNDEFINED_OPERATION, [aItem1.FExpression, aItem1.FOperation, aItem2.FExpression]));
  end;
end;

procedure TLExpression.setNullValueByType(aItem: TLExpression; aType: TItemType);
begin
  aItem.FItemType := aType;
  case aType of
    itInteger: aItem.FValue := 0;
    itFloat: aItem.FValue := 0.00;
    itDate: aItem.FValue := 0;
    itString: aItem.FValue := '';
  end;
end;

procedure TLExpression.setOperation(aOper: String);
begin
  FOperation := aOper;
  FPriority := getOperationPriority(FOperation);
  //Устанавливаем максиамльный приоритет используемых операций
  MaxPriority := Max(MaxPriority, FPriority);
end;

initialization

finalization

end.
