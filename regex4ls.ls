Option Public
Option Declare

UseLSX "*javacon"

%REM Description
	This LotusScript library adds Support for Regular Expressions to LotusScript
	by supplying Wrappers around native Java classes.
	This lets you work with regular LotusScript objects without bothering with
	clumsy LS2J code.

	See the documentation of the individual classes for code examples.
%END REM

REM Version: v1.1.0

REM Author: Harald Albers, HS - Hamburger Software GmbH & Co. KG

%REM Copyright
	Copyright (C) 2010 HS - Hamburger Software GmbH & Co. KG. All rights reserved.

	This library is released under the Apache Licence Version 2.0, see
	    http://www.apache.org/licenses/LICENSE-2.0.txt

	It is distributed in the hope that it will be useful, but without any warranty;
	without even the implied warranty of merchantibility or fitness for a particular purpose.
%END REM

%REM Documentation source
	The classes in this library are wrappers around java.util.regex.Pattern and
	java.util.regex.Matcher. Large portions of these classes' functionality are directly made
	accessable to the user by delegation. For these methods and properties, the documentation
	was taken from the Java documentation, available at
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Matcher.html
%END REM

Public Const RE_DEFAULT = 0

%REM
	java.util.regex.Pattern.UNIX_LINES = 1 [0x1]

	Enables Unix lines mode.

	In this mode, only the '\n' line terminator is recognized in the behavior of ., ^, and $.

	Unix lines mode can also be enabled via the embedded flag expression (?d).
%END REM
Public Const RE_UNIX_LINES = 1

%REM
	java.util.regex.Pattern.CASE_INSENSITIVE = 2 [0x2]

	Enables case-insensitive matching.

	By default, case-insensitive matching assumes that only characters In the US-ASCII charset
	are being matched. Unicode-aware case-insensitive matching can be enabled by specifying the
	RE_UNICODE_CASE flag in conjunction with this flag.

	Case-insensitive matching can also be enabled via the embedded flag expression (?i).
%END REM
Public Const RE_CASE_INSENSITIVE = 2

%REM
	java.util.regex.Pattern.COMMENTS = 4 [0x4]

	Permits whitespace and comments In pattern.

	In this mode, whitespace is ignored, and embedded comments starting with # are ignored
	until the end of a line.

	Comments mode can also be enabled via the embedded flag expression (?x).
%END REM
Public Const RE_COMMENTS = 4

%REM
	java.util.regex.Pattern.MULTILINE = 8 [0x8]

	Enables multiline mode.

	In multiline mode the expressions ^ and $ match just after or just before, respectively,
	a line terminator or the end of the input sequence. By default these expressions only
	match at the beginning and the end of the entire input sequence.

	Multiline mode can also be enabled via the embedded flag expression (?m).
%END REM
Public Const RE_MULTILINE = 8

%REM
	java.util.regex.Pattern.LITERAL = 16 [0x10]

	Enables literal parsing of the pattern.

	When this flag is specified then the input string that specifies the pattern is treated
	as a sequence of literal characters. Metacharacters or escape sequences in the input
	sequence will be given no special meaning.

	The flags RE_CASE_INSENSITIVE and RE_UNICODE_CASE retain their impact on matching when
	used in conjunction with this flag. The other flags become superfluous.

	There is no embedded flag character for enabling literal parsing.
%END REM
Public Const RE_LITERAL = 16

%REM
	java.util.regex.Pattern.DOTALL = 32 [0x20]

	Enables dotall mode.

	In dotall mode, the expression . matches any character, including a line terminator.
	By default this expression does not match line terminators.

	Dotall mode can also be enabled via the embedded flag expression (?s).
	(The s is a mnemonic for "single-line" mode, which is what this is called in Perl.)
%END REM
Public Const RE_DOTALL = 32

%REM
	java.util.regex.Pattern.UNICODE_CASE = 64 [0x40]

	Enables Unicode-aware case folding.

	When this flag is specified then case-insensitive matching, when enabled by the
	RE_CASE_INSENSITIVE flag, is done in a manner consistent with the Unicode Standard.
	By default, case-insensitive matching assumes that only characters in the US-ASCII charset
	are being matched.

	Unicode-aware case folding can also be enabled via the embedded flag expression (?u).
%END REM
Public Const RE_UNICODE_CASE = 64

%REM
	java.util.regex.Pattern.CANON_EQ = 128 [0x80]
	
	Enables canonical equivalence.

	When this flag is specified then two characters will be considered to match if, and only if,
	their full canonical decompositions match. The expression "a\u030A", for example, will match
	the string "\u00E5" when this flag is specified. By default, matching does not take canonical
	equivalence into account.

	There is no embedded flag character for enabling canonical equivalence.
%END REM
Public Const RE_CANON_EQ = 128


%REM
	Base class of the Regular Expression hierarchy. Specifies some common behaviour for
	the classes in this library.
	Do not use, use the derived classes RegexMatcher and SimpleRegexMatcher instead.
%END REM
Class AbstractRegexMatcher

	Private jPatternFactoryMethod As JavaMethod
	Private jMatcher As JavaObject

	Sub New()
		Dim jSession As New JavaSession
		Dim jPatternClass As JavaClass
		Set jPatternClass = jSession.GetClass("java/util/regex/Pattern")
		Set jPatternFactoryMethod = jPatternClass.GetMethod("compile",_
			"(Ljava/lang/String;I)Ljava/util/regex/Pattern;")
	End Sub

	%REM
		Returns a java.util.regex.Matcher object for the given InputText and Pattern with the
		specified Options.
		For valid option values, see the comments in the Declarations section of this library or the
		Field Details at http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html
	%END REM
	Private Function getMatcher(InputText As String, Pattern As String, Options As Integer) As JavaObject
		Dim jPattern As JavaObject
		Set jPattern = jPatternFactoryMethod.Invoke(,Pattern, Options)
		Set getMatcher = jPattern.matcher(InputText)
	End Function

End Class


%REM
	A simple Regular Expression matcher that can determine whether a text contains
	a substring that is described by a Regular Expression (pattern).
	In case of a match, it gives you access to the matching part of the original
	text via the Match property.

	The Regular Expressions used in this class follow java.util.regex.Pattern syntax
	as described in
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html

	To make the matching process case insensitive, either set the IgnoreCase property
	to true or use (?i) in your Regular Expression.

	Example: search for an agent with licence to kill.
		Dim Matcher As New SimpleRegexMatcher
		If Matcher.matches("James Bond is agent 007.", "\b00\d\b") Then
			Print "found agent #" & Matcher.Match
		End If
%END REM
Class SimpleRegexMatcher As AbstractRegexMatcher

	%REM
		Set this property to true if you want the matching process to be case insensitive.
		You can also enable case insensitivity by including the sequence (?i) in your
		Regular Expression. Use (?-i) to disable it again.
	%END REM
	Public IgnoreCase As Boolean

	%REM
		Returns true if the Regular Expression Pattern is found in InputText
	%END REM
	Function matches(InputText As String, Pattern As String) As Boolean
		Set jMatcher = getMatcher(InputText, Pattern, getPatternModifier())
		matches = jMatcher.find()
	End Function

	Private Function getPatternModifier() As Integer
		If IgnoreCase Then
			' the RE_UNICODE_CASE flag extends case sensitivity support to non-ASCII characters
			getPatternModifier = RE_CASE_INSENSITIVE Or RE_UNICODE_CASE
		Else
			getPatternModifier = RE_DEFAULT
		End If
	End Function

	%REM
		Returns the matching part of the last matching process.
		If there was no match, an empty string is returned.
	%END REM
	Property Get Match As String
		If jMatcher Is Nothing Then
			Error 1, TypeName(Me) & ".Match is not available before calling matches()."
		End If
		
		On Error Resume Next
		Match = jMatcher.group()
	End Property

End Class


%REM
	RegexMatcher combines java.util.regex.Pattern and java.util.regex.Matcher into one handy
	LotusScript wrapper to perform Regular Expression operations on text.

	You can do repeated matches on the same string with the find() method to locate multiple
	occurrences of pattern matches, and you can access captured groups with the group() function.
	
	The replaceAll() and replaceFirst() methods let you replace matching parts with strings that
	may contain backreferences to captured parts of the original string.

	In contrast to the SimpleRegexMatcher that is targeted to users new to the java.util.regex
	classes, this class assumes a basic level of familiarity with the underlying Java classes.
	Please see
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Matcher.html
	and
	    http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html
	for behavioural details of the wrapped methods, the pattern syntax and possible exceptions.

	When working with RegexMatcher, you do not use java.util.regex.Pattern objects.
	Instead the text, Regular Expression and flags (matching options that determine aspects like
	case sensitivity and multiline matching) are all specified in the RegexMatcher constructor.

	Example: Find name of someone who introduces himself in a James Bondish way.
	This example features capture groups and a backreference to the first capture group.
		Dim Matcher As New RegexMatcher(_
			"My name is Presley. Elvis Presley.", "name is (\w+)\. (\w+) \1\.", 0)
		If Matcher.find() Then
			Print Matcher.group(2) & " " & Matcher.group(1)  ' prints "Elvis Presley"
		End If

	Example: Replace parts of the input using captured groups
		Dim Matcher As New RegexMatcher("from 10:00 to 12:00", "from (\d\d:\d\d) to (\d\d:\d\d)", 0)
		Print Matcher.replaceAll("$1-$2")  ' prints "10:00-12:00"

	Example: Create case insensitive matcher that also honours case of non-ASCII characters
		Dim Matcher As New RegexMatcher(text$, pattern$, RE_CASE_INSENSITIVE Or RE_UNICODE_CASE)
%END REM
Class RegexMatcher As AbstractRegexMatcher

	%REM
		Creates a RegexMatcher object that works on InputText using the Regular Expression
		Pattern with the given Options.

		Options is an integer value that modifies the behaviour of the RegexMatcher.
		Combine any of the following constants using OR to get a valid option value:
		RE_UNIX_LINES, RE_CASE_INSENSITIVE, RE_COMMENTS, RE_MULTILINE, RE_LITERAL, RE_DOTALL,
		RE_UNICODE_CASE, RE_CANON_EQ.

		For details on these values, see the comments in the Declarations section of this library or
		the Field Details at http://download.oracle.com/javase/6/docs/api/java/util/regex/Pattern.html

		Use 0 or RE_DEFAULT to specify the default behaviour.
	%END REM
	Sub New(InputText As String, Pattern As String, Options As Integer)
		Set jMatcher = getMatcher(InputText, Pattern, Options)
	End Sub

	%REM
		Attempts to find the next subsequence of the input sequence that matches the pattern.
		This method starts at the beginning of the input sequence, or, if a previous invocation
		of the method was successful and the matcher has not since been reset, at the first
		character not matched by the previous match.

		If the match succeeds then the matching part of the input sequence can be obtained via
		the group method.

		Returns true if, and only if, a subsequence of the input sequence matches this
		RegexMatcher's pattern.
	%END REM
	Function find() As Boolean
		find = jMatcher.find()
	End Function

	%REM
		Returns the input subsequence captured by the given group during the previous match
		operation. Specify zero to access the whole match.

		If the pattern contains capturing groups, the portions matched by these can be obtained
		by supplying higher indexes.

		Capturing groups are indexed from left to right, starting at one. Group zero denotes
		the entire pattern.
	%END REM
	Function group(index As Integer) As String
		group = jMatcher.group(index)
	End Function

	%REM
		Returns the number of capturing groups in this RegexMatcher's pattern.
		Group zero denotes the entire pattern by convention. It is not included in this count.
	%END REM
	Property Get GroupCount As Integer
		GroupCount = jMatcher.groupCount()
	End Property

	%REM
		Resets this RegexMatcher. All or its explicit state information is discarded.
	%END REM
	Sub reset()
		Call jMatcher.reset()
	End Sub

	%REM
		Replaces every subsequence of the input sequence that matches the pattern with the given
		replacement string.
	
		This method first resets this RegexMatcher. It then scans the input sequence looking for
		matches of the pattern. Characters that are not part of any match are appended directly
		to the result string; each match is replaced in the result by the replacement string.
		The replacement string may contain backreferences to captured subsequences like $1.
	
		Invoking this method changes this RegexMatcher's state. If the RegexMatcher is to be
		used in further matching operations then it should first be reset.
	%END REM

	Function replaceAll(Replacement As String) As String
		replaceAll = jMatcher.replaceAll(Replacement)
	End Function

	%REM
		Replaces the first subsequence of the input sequence that matches the pattern with the
		given replacement string.

		This method first resets this matcher. It then scans the input sequence looking for a
		match of the pattern. Characters that are not part of the match are appended directly to
		the result string; the match is replaced in the result by the replacement string.
		The replacement string may contain backreferences to captured subsequences like $1.

		Invoking this method changes this RegexMatcher's state. If the RegexMatcher is to be
		used in further matching operations then it should first be reset.
	%END REM
	Function replaceFirst(Replacement As String) As String
		replaceFirst = jMatcher.replaceFirst(Replacement)
	End Function

End Class