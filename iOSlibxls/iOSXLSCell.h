//
//  iOSXLSCell.h
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/19.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import <Foundation/Foundation.h>

typedef enum { cellBlank=0, cellString, cellInteger, cellFloat, cellBool, cellError, cellUnknown } contentsType;

@interface iOSXLSCell : NSObject

@property (nonatomic, assign, readonly) contentsType type;
@property (nonatomic, assign, readonly) uint16_t row;
@property (nonatomic, assign, readonly) char *colStr;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, assign, readonly) uint16_t col;
@property (nonatomic, strong, readonly) NSString *str;        // typeof depends on contentsType
@property (nonatomic, strong, readonly) NSNumber *val;        // typeof depends on contentsType

+ (iOSXLSCell *)blankCell;

// Debugging
- (void)show;
- (NSString *)dump;

@end
