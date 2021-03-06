% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/officerpack.R
\name{lhtext}
\alias{lhtext}
\title{lhtext}
\usage{
lhtext(doc, t, save = NULL, heading = "HD")
}
\arguments{
\item{doc}{doc created by read_docx or lhstartdoc with template. See lhtext_example.}

\item{t}{list of items. see example.}

\item{toc.level}{maximimum toc level}

\item{template}{Word document template could be used for styles. Styles should be mapped in style.to.map. Template is also available at github: to load it, just run  lhtemp() once to download and store the templates in your PC at "c:lhtemplate. Note that the templates and logo are also used in xptdef package.}

\item{style.to.map}{Map the styles in template to be used. Ex: mypar is for footnote (font size)}
}
\description{
Simple way to create word document using loop of from t list.
Type lhtext and copy the template to R workspace and start writing.
}
\keyword{lhtext}
