//
//  QuickExcelReaderUtil.m
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/17.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import "QuickExcelReaderUtil.h"
#import "QuickExcelKitDefine.h"
#import "iOSXLSReader.h"
#import "CSVParser.h"
#import "ZXLSXParser.h"

@interface QuickExcelReaderUtil()<ZXLSXParserDelegate>

@property (nonatomic, strong) ZXLSXParser *xmlPaser;

@property (nonatomic, strong) NSMutableDictionary<NSString*, QuickExcelReaderBlock>*callbacks;
@end

@implementation QuickExcelReaderUtil

static QuickExcelReaderUtil *_sharedUtil = nil;

+ (QuickExcelReaderUtil *) sharedUtil
{
    static dispatch_once_t onceToken;
    dispatch_block_t block = ^{
        if(!_sharedUtil)
        {
            _sharedUtil = [[self class] new];
        }
    };
    if ([NSThread isMainThread])
    {
        dispatch_once(&onceToken, block);
    }
    else
    {
        dispatch_sync(dispatch_get_main_queue(), ^{
            dispatch_once(&onceToken, block);
        });
    }
    return _sharedUtil;
}

-(instancetype)init
{
    if(self = [super init])
    {
        _xmlPaser = [ZXLSXParser defaultZXLSXParser];
        [_xmlPaser setParseOutType:ZParseOutTypeArrayObj];
        _xmlPaser.delegate = self;
        _callbacks = [NSMutableDictionary<NSString*, QuickExcelReaderBlock> dictionary];
    }
    return self;
}

+(void) readExcelWithPath:(NSString*) filePath complete:(QuickExcelReaderBlock)block
{
    NSString *mimeType = [[[filePath componentsSeparatedByString:@"."] lastObject] lowercaseString];
    
    if([mimeType isEqualToString:@"csv"])
    {
        [[QuickExcelReaderUtil sharedUtil]  parserExcel_CSV_WithPath:filePath complete:block];
    }
    else if ([mimeType isEqualToString:@"xls"])
    {
        [[QuickExcelReaderUtil sharedUtil] parserExcel_XLS_WithPath:filePath complete:block];
    }
    else if ([mimeType isEqualToString:@"xlsx"])
    {
        [[QuickExcelReaderUtil sharedUtil] parserExcel_XLSX_WithPath:filePath complete:block];
    }
    else
    {
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel 失败，格式 %@ 不支持。", mimeType];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
    }
}

-(void)parserExcel_XLSX_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block
{
    if(block)
    {
        [self.callbacks removeAllObjects];
        [self.callbacks setObject:block forKey:filePath];
        [_xmlPaser setParseFilePath:filePath];
        _xmlPaser.delegate = self;
        [_xmlPaser parse];
    }
}

-(void)parserExcel_XLS_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block
{
    iOSXLSReader *reader = [iOSXLSReader xlsReaderWithPath:filePath];
    if(![reader isKindOfClass:[iOSXLSReader class]])
    {
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        return;
    }
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    uint32_t sheetsCount = [reader numberOfSheets];
    for (uint32_t i = 0; i < sheetsCount; i++)
    {
        NSString *sheetName = [reader sheetNameAtIndex:i];
        NSMutableArray<ZContent *> *resultArray = [NSMutableArray array];
        [reader startIterator:i];
        int rows = [reader numberOfRowsInSheet:i];
        int cols = [reader numberOfColsInSheet:i];
        for(int r = 1; r <= rows; r++)
        {
            for(int c = 1;c <= cols; c++)
            {
                unichar ch =64 + c;
                NSString *str =[NSString stringWithUTF8String:(char *)&ch];
                iOSXLSCell *cell = [reader cellInWorkSheetIndex:i row:r col:c];
                ZContent *content = [[ZContent alloc] init];
                content.sheetName = [reader sheetNameAtIndex:i];
                content.keyName = [NSString stringWithFormat: @"%@%d", str,c+r-1];
                content.value = [cell dump];
                [resultArray addObject:content];
            }
        }
        [results setObject:resultArray forKey:sheetName];
    }
    if(block)block(results, nil);
}

-(void)parserExcel_CSV_WithPath:(NSString *)filePath complete:(QuickExcelReaderBlock)block
{
    NSMutableArray *array = [CSVParser readCSVData:filePath];
    if(![array isKindOfClass:[NSArray class]])
    {
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        return;
    }
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    NSString *sheetName = @"csv";
    NSMutableArray<ZContent *> *resultArray = [NSMutableArray array];
    int row = 0;
    for(NSArray *item in array)
    {
        for(int i = 0; i < item.count; i++) {
            
            unichar ch =65 + i;
            NSString *str =[NSString stringWithUTF8String:(char *)&ch];
            ZContent *content = [[ZContent alloc] init];
            content.sheetName = sheetName;
            content.keyName = [NSString stringWithFormat: @"%@%d", str,i+1+row];
            content.value = item[i];
            [resultArray addObject:content];
        }
        row++;
    }
    [results setObject:resultArray forKey:sheetName];
    if(block)block(results, nil);
}


#pragma -mark ZXLSXParserDelegate
- (void)parser:(ZXLSXParser *)parser success:(id)responseObj
{
    NSString *filePath = parser.parseFilePath;
    QuickExcelReaderBlock block = [self.callbacks objectForKey:filePath];
    if(![responseObj isKindOfClass:[NSArray<ZContent *> class]])
    {
        NSString *errorString = [NSString stringWithFormat:@"读取 Excel: %@ 失败。", [filePath lastPathComponent]];
        if(block)block(nil, QUICKEXCELKIT_ERROR(1000, errorString));
        [self.callbacks removeAllObjects];
        return;
    }
    NSMutableDictionary<NSString*, NSArray<ZContent*>*>* results = [NSMutableDictionary<NSString*, NSArray<ZContent*>*> dictionary];
    NSString *defaultSheetName = @"xlsx";
    NSString *defaultKeyName = @"xlsx_key";
    for (ZContent * content in responseObj) {
        if(!content.sheetName) content.sheetName = defaultSheetName;
        if(!content.keyName) content.keyName = defaultKeyName;
        NSMutableArray<ZContent *> *resultArray = (NSMutableArray<ZContent *> *)[results objectForKey:content.sheetName];
        if(![resultArray isKindOfClass:[NSMutableArray<ZContent *> class]])
        {
            resultArray = [NSMutableArray array];
            [results setObject:resultArray forKey:content.sheetName];
        }
        [resultArray addObject:content];
    }
    if(block)block(results, nil);
    [self.callbacks removeAllObjects];
}
@end
