class Header():

    def fte():
        return r"""
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endfirsthead
     
        \multicolumn{6}{c}{\textit{
            \textnormal{\textbf{\shortstack[c]{
                CITY OF DETROIT%
                \\ {BUDGET DEVELOPMENT}%
                \\ {POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER}%
                \\ {DEPARTMENT 72 - DETROIT PUBLIC LIBRARY}}
            }
            \rule[-1em]{0pt}{6em} % Bottom space adjustment
        }}} \\        
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endhead"""

    def summary():
        return r"""\begin{center}%
\textbf{%
CITY OF DETROIT%
\\ {BUDGET DEVELOPMENT}%
\\ {POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER}%
\\ {DEPARTMENT 72 - DETROIT PUBLIC LIBRARY}}%
\end{center}%

            \renewcommand{\arraystretch}{1.3} % Increase the row height
            \setlength{\tabcolsep}{4pt}       % Add padding
        %
\arrayrulecolor{black}\begin{longtable}{>{\arraybackslash}p{11.5cm}>{\centering\arraybackslash}p{2.25cm}>{\centering\arraybackslash}p{2.25cm}>{\centering\arraybackslash}p{2.25cm}>{\centering\arraybackslash}p{2.25cm}>{\centering\arraybackslash}p{2.25cm}}

        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endfirsthead
     
        \multicolumn{6}{c}{\textit{
            \textnormal{\textbf{\shortstack[c]{
                CITY OF DETROIT%
                \\ {BUDGET DEVELOPMENT}%
                \\ {POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER}%
                \\ {DEPARTMENT 72 - DETROIT PUBLIC LIBRARY}}
            }
            \rule[-1em]{0pt}{6em} % Bottom space adjustment
        }}} \\        
        \specialrule{1.5pt}{0pt}{0pt}
        \multicolumn{1}{|l}{
        \textbf{\shortstack[l]{
        \rule{0pt}{2em} % Top space adjustment
        Department \# - Department Name\\
        \hspace{0.5cm}Fund \# - Fund Name\\
        \hspace{1cm}Appropriation \# - Appropriation Name\\
        \hspace{1.5cm}Cost Center \# - Cost Center Name\\
        \hspace{2cm}Job Code \# - Job Title}}
        \rule[-1em]{0pt}{4em} % Bottom space adjustment
        } &
        \textbf{\shortstack{FY2025 \\ Adopted}} &
        \textbf{\shortstack{FY2026 \\ Adopted}} &
        \textbf{\shortstack{FY2027 \\ Forecast}} &
        \textbf{\shortstack{FY2028 \\ Forecast}} &
        \multicolumn{1}{c|}{
        \textbf{\shortstack{FY2029 \\ Forecast}}
        } \\
        \specialrule{1.5pt}{0pt}{0pt}

\endhead"""

    def text(main, subheaders, column_format, 
             fund,
             category,
             approp,
             cc,
             jobName):
        
        top_lines = rf"""
            {main} \\
            {subheaders[0]} \\
            {subheaders[1]} \\
            {subheaders[2]}
        """

        top =  rf"""
            \begin{{center}}%
            \textbf{{{top_lines}
            }}
            \end{{center}}%

                        \renewcommand{{\arraystretch}}{{1.3}} % Increase the row height
                        \setlength{{\tabcolsep}}{{4pt}}       % Add padding
                    %
            \arrayrulecolor{{black}}
            \begin{{longtable}} {{{column_format}}}
                    \specialrule{{1.5pt}}{{0pt}}{{0pt}}
                    \multicolumn{{1}}{{|l}}{{"""
            
        title = r"""
            \textbf{{\shortstack[l]{{
                \rule{{0pt}}{{2em}} % Top space adjustment
                Department \# - Department Name\\"""
        
        if fund:
            title += r'\hspace{{0.5cm}}Fund \# - Fund Name\\'

        if category:
            title += r'\hspace{{1cm}}Summary Category\\'

        if approp:
            title += r'\hspace{{1cm}}Appropriation \# - Appropriation Name\\'

        if cc:
            title += r'\hspace{{1.5cm}}Cost Center \# - Cost Center Name\\'

        if jobName:
            title += r'\hspace{{2cm}}Job Code \# - Job Title'

        title += r"""}}
            \rule[-1em]{{0pt}}{{4em}} % Bottom space adjustment
            } &
            \textbf{\shortstack{FY2025 \\ Adopted}} &
            \textbf{\shortstack{FY2026 \\ Adopted}} &
            \textbf{\shortstack{FY2027 \\ Forecast}} &
            \textbf{\shortstack{FY2028 \\ Forecast}} &
            \multicolumn{1}{c|}{
            \textbf{\shortstack{FY2029 \\ Forecast}}
            } \\
            \specialrule{1.5pt}{0pt}{0pt}
            """
        
        firsthead = top + title + '\n' + r'\endfirsthead'
        
        head = r"""
            \multicolumn{6}{c}{\textit{
                \textnormal{\textbf{\shortstack[c]{
                    CITY OF DETROIT%
                    \\ {BUDGET DEVELOPMENT}%
                    \\ {POSITION DETAIL BY DEPARTMENT, FUND, APPROPRIATION, \& COST CENTER}%
                    \\ {DEPARTMENT 72 - DETROIT PUBLIC LIBRARY}}
                }
                \rule[-1em]{0pt}{6em} % Bottom space adjustment
            }}} \\        
            \specialrule{1.5pt}{0pt}{0pt}
            \multicolumn{1}{|l}{
            \textbf{\shortstack[l]{
            \rule{0pt}{2em} % Top space adjustment
            Department \# - Department Name\\
            \hspace{0.5cm}Fund \# - Fund Name\\
            \hspace{1cm}Appropriation \# - Appropriation Name\\
            \hspace{1.5cm}Cost Center \# - Cost Center Name\\
            \hspace{2cm}Job Code \# - Job Title}}
            \rule[-1em]{0pt}{4em} % Bottom space adjustment
            } &
            \textbf{\shortstack{FY2025 \\ Adopted}} &
            \textbf{\shortstack{FY2026 \\ Adopted}} &
            \textbf{\shortstack{FY2027 \\ Forecast}} &
            \textbf{\shortstack{FY2028 \\ Forecast}} &
            \multicolumn{1}{c|}{
            \textbf{\shortstack{FY2029 \\ Forecast}}
            } \\
            \specialrule{1.5pt}{0pt}{0pt}

            \endhead

            """
        return firsthead + head