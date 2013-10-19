# coding=UTF-8
# Copyright (c) 2010-2011 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

from numbers import Number

from openpyxl.shared.xmltools import Element, SubElement, get_document_content
from openpyxl.shared.compat.itertools import iteritems
from openpyxl.chart import Chart, ErrorBar, BarChart, LineChart, PieChart, ScatterChart

try:
    # Python 2
    basestring
except NameError:
    # Python 3
    basestring = str


def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, Number):
        value = "%.15g" % value
    elif not isinstance(value, basestring):
        value = str(value)
    return value

class BaseChartWriter(object):

    def __init__(self, chart):
        self.chart = chart

    def write(self):
        """ write a chart """

        root = Element('c:chartSpace',
            {'xmlns:c':"http://schemas.openxmlformats.org/drawingml/2006/chart",
             'xmlns:a':"http://schemas.openxmlformats.org/drawingml/2006/main",
             'xmlns:r':"http://schemas.openxmlformats.org/officeDocument/2006/relationships"})

        SubElement(root, 'c:lang', {'val':self.chart.lang})
        self._write_chart(root)
        self._write_print_settings(root)
        self._write_shapes(root)

        return get_document_content(root)

    def _write_title(self, chart):
        if self.chart.title != '':
            title = SubElement(chart, 'c:title')
            tx = SubElement(title, 'c:tx')
            rich = SubElement(tx, 'c:rich')
            SubElement(rich, 'a:bodyPr')
            SubElement(rich, 'a:lstStyle')
            p = SubElement(rich, 'a:p')
            pPr = SubElement(p, 'a:pPr')
            SubElement(pPr, 'a:defRPr')
            r = SubElement(p, 'a:r')
            SubElement(r, 'a:rPr', {'lang':self.chart.lang})
            t = SubElement(r, 'a:t').text = self.chart.title
            SubElement(title, 'c:layout')

    def _write_axis(self, plot_area, axis, label):
        if self.chart.type == Chart.PIE_CHART:
            return()

        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, 'c:axId', {'val':safe_string(axis.id)})

        scaling = SubElement(ax, 'c:scaling')
        SubElement(scaling, 'c:orientation', {'val':axis.orientation})
        if label == 'c:valAx':
            SubElement(scaling, 'c:max', {'val':str(float(axis.max))})
            SubElement(scaling, 'c:min', {'val':str(float(axis.min))})

        SubElement(ax, 'c:axPos', {'val':axis.position})
        if label == 'c:valAx':
            SubElement(ax, 'c:majorGridlines')
            SubElement(ax, 'c:numFmt', {'formatCode':"General", 'sourceLinked':'1'})
        if axis.title != '':
            title = SubElement(ax, 'c:title')
            tx = SubElement(title, 'c:tx')
            rich = SubElement(tx, 'c:rich')
            SubElement(rich, 'a:bodyPr')
            SubElement(rich, 'a:lstStyle')
            p = SubElement(rich, 'a:p')
            pPr = SubElement(p, 'a:pPr')
            SubElement(pPr, 'a:defRPr')
            r = SubElement(p, 'a:r')
            SubElement(r, 'a:rPr', {'lang':self.chart.lang})
            t = SubElement(r, 'a:t').text = axis.title
            SubElement(title, 'c:layout')
        SubElement(ax, 'c:tickLblPos', {'val':axis.tick_label_position})
        SubElement(ax, 'c:crossAx', {'val':str(axis.cross)})
        SubElement(ax, 'c:crosses', {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, 'c:auto', {'val':'1'})
        if axis.label_align:
            SubElement(ax, 'c:lblAlgn', {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, 'c:lblOffset', {'val':str(axis.label_offset)})
        if label == 'c:valAx':
            if self.chart.type == Chart.SCATTER_CHART:
                SubElement(ax, 'c:crossBetween', {'val':'midCat'})
            else:
                SubElement(ax, 'c:crossBetween', {'val':'between'})
            SubElement(ax, 'c:majorUnit', {'val':str(float(axis.unit))})

    def _write_series(self, subchart):

        for i, serie in enumerate(self.chart._series):
            ser = SubElement(subchart, 'c:ser')
            SubElement(ser, 'c:idx', {'val':safe_string(i)})
            SubElement(ser, 'c:order', {'val':safe_string(i)})

            if serie.legend:
                tx = SubElement(ser, 'c:tx')
                self._write_serial(tx, serie.legend)

            if serie.color:
                sppr = SubElement(ser, 'c:spPr')
                if self.chart.type == Chart.BAR_CHART:
                    # fill color
                    fillc = SubElement(sppr, 'a:solidFill')
                    SubElement(fillc, 'a:srgbClr', {'val':serie.color})
                # edge color
                ln = SubElement(sppr, 'a:ln')
                fill = SubElement(ln, 'a:solidFill')
                SubElement(fill, 'a:srgbClr', {'val':serie.color})

            if serie.error_bar:
                self._write_error_bar(ser, serie)

            if serie.labels:
                cat = SubElement(ser, 'c:cat')
                self._write_serial(cat, serie.labels)

            if self.chart.type == Chart.SCATTER_CHART:
                if serie.xvalues:
                    xval = SubElement(ser, 'c:xVal')
                    self._write_serial(xval, serie.xreference)

                val = SubElement(ser, 'c:yVal')
            else:
                val = SubElement(ser, 'c:val')
            self._write_serial(val, serie.reference)

    def _write_serial(self, node, reference, literal=False):

        is_ref = hasattr(reference, 'pos1')
        data_type = reference.data_type
        number_format = getattr(reference, 'number_format')

        mapping = {'n':{'ref':'numRef', 'cache':'numCache'},
                   's':{'ref':'strRef', 'cache':'strCache'}}

        if is_ref:
            ref = SubElement(node, 'c:%s' % mapping[data_type]['ref'])
            SubElement(ref, 'c:f').text = str(reference)
            data = SubElement(ref, 'c:%s' % mapping[data_type]['cache'])
            values = reference.values
        else:
            data = SubElement(node, 'c:numLit')
            values = (1,)

        if data_type == 'n':
            SubElement(data, 'c:formatCode').text = number_format or 'General'

        SubElement(data, 'c:ptCount', {'val':str(len(values))})
        for j, val in enumerate(values):
            point = SubElement(data, 'c:pt', {'idx':str(j)})
            if not isinstance(val, basestring):
                val = safe_string(val)
            SubElement(point, 'c:v').text = val

    def _write_error_bar(self, node, serie):

        flag = {ErrorBar.PLUS_MINUS:'both',
                ErrorBar.PLUS:'plus',
                ErrorBar.MINUS:'minus'}

        eb = SubElement(node, 'c:errBars')
        SubElement(eb, 'c:errBarType', {'val':flag[serie.error_bar.type]})
        SubElement(eb, 'c:errValType', {'val':'cust'})

        plus = SubElement(eb, 'c:plus')
        # apart from setting the type of data series the following has
        # no effect on the writer
        self._write_serial(plus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.MINUS))

        minus = SubElement(eb, 'c:minus')
        self._write_serial(minus, serie.error_bar.reference,
            literal=(serie.error_bar.type == ErrorBar.PLUS))

    def _write_legend(self, chart):

        if self.chart.show_legend:
            legend = SubElement(chart, 'c:legend')
            SubElement(legend, 'c:legendPos', {'val':self.chart.legend.position})
            SubElement(legend, 'c:layout')

    def _write_print_settings(self, root):

        settings = SubElement(root, 'c:printSettings')
        SubElement(settings, 'c:headerFooter')
        try:
            # Python 2
            print_margins_items = iteritems(self.chart.print_margins)
        except AttributeError:
            # Python 3
            print_margins_items = self.chart.print_margins.items()

        margins = dict([(k, safe_string(v)) for (k, v) in print_margins_items])
        SubElement(settings, 'c:pageMargins', margins)
        SubElement(settings, 'c:pageSetup')

    def _write_shapes(self, root):

        if self.chart._shapes:
            SubElement(root, 'c:userShapes', {'r:id':'rId1'})

    def write_rels(self, drawing_id):

        root = Element('Relationships', {'xmlns' : 'http://schemas.openxmlformats.org/package/2006/relationships'})
        attrs = {'Id' : 'rId1',
            'Type' : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes',
            'Target' : '../drawings/drawing%s.xml' % drawing_id }
        SubElement(root, 'Relationship', attrs)
        return get_document_content(root)


class PieChartWriter(BaseChartWriter):

    def _write_chart(self, root):
        chart = self.chart

        ch = SubElement(root, 'c:chart')
        self._write_title(ch)
        plot_area = SubElement(ch, 'c:plotArea')
        layout = SubElement(plot_area, 'c:layout')
        mlayout = SubElement(layout, 'c:manualLayout')
        SubElement(mlayout, 'c:layoutTarget', {'val':'inner'})
        SubElement(mlayout, 'c:xMode', {'val':'edge'})
        SubElement(mlayout, 'c:yMode', {'val':'edge'})
        SubElement(mlayout, 'c:x', {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, 'c:y', {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, 'c:w', {'val':safe_string(chart.width)})
        SubElement(mlayout, 'c:h', {'val':safe_string(chart.height)})

        subchart = SubElement(plot_area, 'c:pieChart')
        SubElement(subchart, 'c:varyColors', {'val':'1'})

        self._write_series(subchart)

        self._write_legend(ch)

        SubElement(ch, 'c:plotVisOnly', {'val':'1'})


class LineChartWriter(BaseChartWriter):

    def _write_chart(self, root):

        chart = self.chart

        ch = SubElement(root, 'c:chart')
        self._write_title(ch)
        plot_area = SubElement(ch, 'c:plotArea')
        layout = SubElement(plot_area, 'c:layout')
        mlayout = SubElement(layout, 'c:manualLayout')
        SubElement(mlayout, 'c:layoutTarget', {'val':'inner'})
        SubElement(mlayout, 'c:xMode', {'val':'edge'})
        SubElement(mlayout, 'c:yMode', {'val':'edge'})
        SubElement(mlayout, 'c:x', {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, 'c:y', {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, 'c:w', {'val':safe_string(chart.width)})
        SubElement(mlayout, 'c:h', {'val':safe_string(chart.height)})

        subchart = SubElement(plot_area, 'c:lineChart')

        SubElement(subchart, 'c:grouping', {'val':chart.grouping})

        self._write_series(subchart)

        SubElement(subchart, 'c:axId', {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, 'c:axId', {'val':safe_string(chart.y_axis.id)})

        self._write_axis(plot_area, chart.x_axis, 'c:catAx')
        self._write_axis(plot_area, chart.y_axis, 'c:valAx')

        self._write_legend(ch)

        SubElement(ch, 'c:plotVisOnly', {'val':'1'})


class BarChartWriter(BaseChartWriter):

    def _write_chart(self, root):

        chart = self.chart

        ch = SubElement(root, 'c:chart')
        self._write_title(ch)
        plot_area = SubElement(ch, 'c:plotArea')
        layout = SubElement(plot_area, 'c:layout')
        mlayout = SubElement(layout, 'c:manualLayout')
        SubElement(mlayout, 'c:layoutTarget', {'val':'inner'})
        SubElement(mlayout, 'c:xMode', {'val':'edge'})
        SubElement(mlayout, 'c:yMode', {'val':'edge'})
        SubElement(mlayout, 'c:x', {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, 'c:y', {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, 'c:w', {'val':safe_string(chart.width)})
        SubElement(mlayout, 'c:h', {'val':safe_string(chart.height)})

        subchart = SubElement(plot_area, 'c:barChart')
        SubElement(subchart, 'c:barDir', {'val':'col'})

        SubElement(subchart, 'c:grouping', {'val':chart.grouping})

        self._write_series(subchart)

        SubElement(subchart, 'c:axId', {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, 'c:axId', {'val':safe_string(chart.y_axis.id)})

        self._write_axis(plot_area, chart.x_axis, 'c:catAx')
        self._write_axis(plot_area, chart.y_axis, 'c:valAx')

        self._write_legend(ch)

        SubElement(ch, 'c:plotVisOnly', {'val':'1'})


class ScatterChartWriter(BaseChartWriter):

    def _write_chart(self, root):

        chart = self.chart

        ch = SubElement(root, 'c:chart')
        self._write_title(ch)
        plot_area = SubElement(ch, 'c:plotArea')
        layout = SubElement(plot_area, 'c:layout')
        mlayout = SubElement(layout, 'c:manualLayout')
        SubElement(mlayout, 'c:layoutTarget', {'val':'inner'})
        SubElement(mlayout, 'c:xMode', {'val':'edge'})
        SubElement(mlayout, 'c:yMode', {'val':'edge'})
        SubElement(mlayout, 'c:x', {'val':safe_string(chart.margin_left)})
        SubElement(mlayout, 'c:y', {'val':safe_string(chart.margin_top)})
        SubElement(mlayout, 'c:w', {'val':safe_string(chart.width)})
        SubElement(mlayout, 'c:h', {'val':safe_string(chart.height)})

        subchart = SubElement(plot_area, 'c:scatterChart')
        SubElement(subchart, 'c:scatterStyle', {'val':'lineMarker'})

        self._write_series(subchart)
        SubElement(subchart, 'c:axId', {'val':safe_string(chart.x_axis.id)})
        SubElement(subchart, 'c:axId', {'val':safe_string(chart.y_axis.id)})

        self._write_axis(plot_area, chart.x_axis, 'c:valAx')
        self._write_axis(plot_area, chart.y_axis, 'c:valAx')

        self._write_legend(ch)

        SubElement(ch, 'c:plotVisOnly', {'val':'1'})

    def _write_axis(self, plot_area, axis, label):
        self.chart.compute_axes()

        ax = SubElement(plot_area, label)
        SubElement(ax, 'c:axId', {'val':safe_string(axis.id)})

        scaling = SubElement(ax, 'c:scaling')
        SubElement(scaling, 'c:orientation', {'val':axis.orientation})
        if label == 'c:valAx':
            SubElement(scaling, 'c:max', {'val':str(float(axis.max))})
            SubElement(scaling, 'c:min', {'val':str(float(axis.min))})

        SubElement(ax, 'c:axPos', {'val':axis.position})
        if label == 'c:valAx':
            SubElement(ax, 'c:majorGridlines')
            SubElement(ax, 'c:numFmt', {'formatCode':"General", 'sourceLinked':'1'})
        if axis.title != '':
            title = SubElement(ax, 'c:title')
            tx = SubElement(title, 'c:tx')
            rich = SubElement(tx, 'c:rich')
            SubElement(rich, 'a:bodyPr')
            SubElement(rich, 'a:lstStyle')
            p = SubElement(rich, 'a:p')
            pPr = SubElement(p, 'a:pPr')
            SubElement(pPr, 'a:defRPr')
            r = SubElement(p, 'a:r')
            SubElement(r, 'a:rPr', {'lang':self.chart.lang})
            t = SubElement(r, 'a:t').text = axis.title
            SubElement(title, 'c:layout')
        SubElement(ax, 'c:tickLblPos', {'val':axis.tick_label_position})
        SubElement(ax, 'c:crossAx', {'val':str(axis.cross)})
        SubElement(ax, 'c:crosses', {'val':axis.crosses})
        if axis.auto:
            SubElement(ax, 'c:auto', {'val':'1'})
        if axis.label_align:
            SubElement(ax, 'c:lblAlgn', {'val':axis.label_align})
        if axis.label_offset:
            SubElement(ax, 'c:lblOffset', {'val':str(axis.label_offset)})
        if label == 'c:valAx':
            SubElement(ax, 'c:crossBetween', {'val':'midCat'})
            SubElement(ax, 'c:majorUnit', {'val':str(float(axis.unit))})


class ChartWriter(object):
    """
    Preserve interface for chart writer
    """

    def __init__(self, chart):
        if isinstance(chart, PieChart):
            self.cw = PieChartWriter(chart)
        elif isinstance(chart, LineChart):
            self.cw = LineChartWriter(chart)
        elif isinstance(chart, BarChart):
            self.cw = BarChartWriter(chart)
        elif isinstance(chart, ScatterChart):
            self.cw = ScatterChartWriter(chart)
        else:
            raise ValueError("Don't know how to handle %s", chart.__class__.__name__)

    def write(self):
        return self.cw.write()
