//
//  QuickExcelReaderUtil.h
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/17.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "ZContent.h"

typedef void (^QuickExcelReaderBlock)(NSDictionary<NSString*, NSArray<ZContent*>*>* results, NSError* error);

@interface QuickExcelReaderUtil : NSObject

+(void) readExcelWithPath:(NSString*) filePath complete:(QuickExcelReaderBlock)block;

@end
