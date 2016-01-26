# regex4ls - Regular Expressions for LotusScript

## OPENNTF

This project is an [OpenNTF](https://openntf.org/) project, and is available under the Apache Licence V2.0. All other aspects of the project, including contributions, defect reports, discussions, feature requests and reviews are subject to the OpenNTF Terms of Use - available at [http://openntf.org/Internal/home.nsf/dx/Terms_of_Use](http://openntf.org/Internal/home.nsf/dx/Terms_of_Use).

[Project page on OpenNTF](https://openntf.org/main.nsf/project.xsp?r=project/Regular%20Expressions%20for%20LotusScript/)

## About

This project enables LotusScript developers to use Regular Expressions in their code. 
It does so by providing a script library with simple wrappers around Java classes that are invoked by [LS2J](https://www-01.ibm.com/support/knowledgecenter/SSVRGU_9.0.1/com.ibm.designer.domino.main.doc/LSAZ_ABOUT_LS2J.html). 
The wrapper classes behave like native LotusScript classes. All the clumsiness of JNI declarations that comes with using LS2J is abstracted away from the user.

The `SimpleRegexMatcher` can determine whether a text matches a pattern. 
In case of a match it gives you access to the matching part of the original text via its `Match` property.

The `RegexMatcher` combines `java.util.regex.Pattern` and `java.util.regex.Matcher` into one handy LotusScript wrapper to perform Regular Expression operations on text. 
You can do repeated matches on the same string with `find()` to locate multiple occurrences of pattern matches, and you can access captured groups with `group()`. 
The `replaceAll()` and `replaceFirst()` methods let you replace matching parts with strings that may contain backreferences to captured parts of the original string. 
Whereas the `SimpleRegexMatcher` is targeted to users that are new to the `java.util.regex` classes, `RegexMatcher` assumes a basic level of familiarity with the underlying Java classes.

The Regular Expressions used in this class follow `java.util.regex.Pattern` syntax as described [here](http://docs.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html).

Intended audience: LotusScript developers.

## Examples

These examples assume that you [installed](#how-to-install) the _regex4ls_ library and included it in the current code unit with
```vbnet
Use "Regular Expressions"
```

This example uses the `SimpleRegexMatcher` to test for a match. The matching part is then printed.
```vbnet
Dim Matcher As New SimpleRegexMatcher
' Search for 00 Agents in a string that contains "007". \d matches on word boundaries.
If Matcher.matches("James Bond is agent 007.", "\b00[1-9]\b") Then
    ' This prints "found agent #007"
    Print "found agent #" & Matcher.match
End If
```

This example features capture groups and a backreference to the first capture group.
```vbnet
' Search for people who introduce themselves in a James Bondish way
Dim Matcher As New RegexMatcher("My name is Presley. Elvis Presley.", "name is (\w+)\. (\w+) \1\.", 0)
If Matcher.find() Then
    ' This prints "Elvis Presley"
    Print Matcher.group(2) & " " & Matcher.group(1)
End If
```

Replace parts of the input using capture groups:
```vbnet
Dim Matcher As New RegexMatcher("From 10:00 To 12:00", "From (\d\d:\d\d) To (\d\d:\d\d)", 0)
' This prints "10:00-12:00"
Print Matcher.replaceAll("$1-$2")
```

## How to Install

Download the [latest release](https://github.com/OpenNTF/regex4ls/releases/latest). Extract _regex4ls.ls_ from the zip file.

In Domino Designer, create a LotusScript library with the name "Regular Expressions". Copy the contents of the downloaded file into the library. Save. 

In your client code (agent, another script library, ...), include the library with `Use "Regular Expressions"`.