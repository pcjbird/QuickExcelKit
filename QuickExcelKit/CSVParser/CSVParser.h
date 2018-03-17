//
//  CSVParser.h
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/17.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface CSVParser : NSObject

+ (NSMutableArray *)readCSVData:(NSString *)filePath;

@end
