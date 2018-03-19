//
//  iOSXLSReader.h
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/19.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "iOSXLSCell.h"

enum {iOSWorkSheetNotFound = UINT32_MAX};

@interface iOSXLSReader : NSObject

+ (iOSXLSReader *)xlsReaderWithPath:(NSString *)filePath;

- (NSString *)libaryVersion;

// Sheet Information
- (uint32_t)numberOfSheets;
- (NSString *)sheetNameAtIndex:(uint32_t)index;
- (uint16_t)rowsForSheetAtIndex:(uint32_t)idx;
- (BOOL)isSheetVisibleAtIndex:(NSUInteger)index;
- (uint16_t)numberOfRowsInSheet:(uint32_t)sheetIndex;
- (uint16_t)numberOfColsInSheet:(uint32_t)sheetIndex;

// Random Access
- (iOSXLSCell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row col:(uint16_t)col;        // uses 1 based indexing!
- (iOSXLSCell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row colStr:(char *)col;        // "A"...."Z" "AA"..."ZZ"

// Iterate through all cells
- (void)startIterator:(uint32_t)sheetNum;
- (iOSXLSCell *)nextCell;

// Summary Information
- (NSString *)appName;
- (NSString *)author;
- (NSString *)category;
- (NSString *)comment;
- (NSString *)company;
- (NSString *)keywords;
- (NSString *)lastAuthor;
- (NSString *)manager;
- (NSString *)subject;
- (NSString *)title;


@end
