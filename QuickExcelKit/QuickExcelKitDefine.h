//
//  QuickExcelKitDefine.h
//  QuickExcelKit
//
//  Created by pcjbird on 2018/3/17.
//  Copyright © 2018年 Zero Status. All rights reserved.
//

#ifndef QuickExcelKitDefine_h
#define QuickExcelKitDefine_h

#define QUICKEXCELKIT_ERROR(ecode, msg)  [NSError errorWithDomain:@"QuickExcelKit" code:(ecode) userInfo:([NSDictionary dictionaryWithObjectsAndKeys:(msg), @"message", nil])]

#endif /* QuickExcelKitDefine_h */
