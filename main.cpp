#include <QApplication>
#include <QtConcurrentRun>
#include <shlobj.h>
#include <QObject>
#include <QFile>
#include <QAxBase>
#include <QAxObject>
#include <QFileDialog>
#include <QScopedPointer>
#include <QDebug>
#include <QString>
#include <QVector>
#include <QList>
#include <QThread>
#include <QMutex>
#include <QMap>
#include <QSet>
#include <QDateTime>
#include <QTextStream>
#include <assert.h>

QVariant MyFunction(QString _filename, int _maxCount)
{
    QString currTime=QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    qDebug() << "MyFunction "
             << QThread::currentThread()->currentThreadId()
             << currTime;

    qDebug("TestConnector: Initializing COM library");
    HRESULT h_result = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    switch(h_result)
    {
    case S_OK:
        qDebug("TestConnector: The COM library was initialized successfully on this thread");
        break;
    case S_FALSE:
        qWarning("TestConnector: The COM library is already initialized on this thread");
        break;
    case RPC_E_CHANGED_MODE:
        qWarning() << "TestConnector: A previous call to CoInitializeEx specified the concurrency model for this thread as multithread apartment (MTA)."
                   << " This could also indicate that a change from neutral-threaded apartment to single-threaded apartment has occurred";
        break;
    }

    // получаем указатель на Excel
    QScopedPointer<QAxObject> excel(new QAxObject("Excel.Application"));
    if(excel.isNull())
    {
        QString error="Cannot get Excel.Application";
        return QVariant(error);
    }

    QScopedPointer<QAxObject> workbooks(excel->querySubObject("Workbooks"));
    if(workbooks.isNull())
    {
        QString error="Cannot query Workbooks";
        return QVariant(error);
    }

    // на директорию, откуда грузить книгу
    QScopedPointer<QAxObject> workbook(workbooks->querySubObject(
                                           "Open(const QString&)",
                                           _filename)
                                       );
    if(workbook.isNull())
    {
        QString error=
                QString("Cannot query workbook.Open(const %1)")
                .arg(_filename);
        return QVariant(error);
    }

    QScopedPointer<QAxObject> sheets(workbook->querySubObject("Sheets"));
    if(sheets.isNull())
    {
        QString error="Cannot query Sheets";
        return QVariant(error);
    }

    int count = sheets->dynamicCall("Count()").toInt(); //получаем кол-во листов
    QStringList readedSheetNames;
    //читаем имена листов
    for (int i=1; i<=count; i++)
    {
        QScopedPointer<QAxObject> sheetItem(sheets->querySubObject("Item(int)", i));
        if(sheetItem.isNull())
        {
            QString error="Cannot query Item(int)"+QString::number(i);
            return QVariant(error);
        }
        readedSheetNames.append( sheetItem->dynamicCall("Name()").toString() );
        sheetItem->clear();
    }
    qDebug() << "XlsWorker readedSheetNames" << readedSheetNames;
    // проходим по всем листам документа
    int sheetNumber=0;
    QMap<QString, QVariant> data;
    foreach (QString sheetName, readedSheetNames)
    {
        QScopedPointer<QAxObject> sheet(
                    sheets->querySubObject("Item(const QVariant&)",
                                           QVariant(sheetName))
                    );
        if(sheet.isNull())
        {
            QString error=
                    QString("Cannot query Item(const %1)")
                    .arg(sheetName);
            return QVariant(error);
        }

        QScopedPointer<QAxObject> usedRange(sheet->querySubObject("UsedRange"));
        QScopedPointer<QAxObject> usedRows(usedRange->querySubObject("Rows"));
        QScopedPointer<QAxObject> usedCols(usedRange->querySubObject("Columns"));
        int rows = usedRows->property("Count").toInt();
        int cols = usedCols->property("Count").toInt();

        //если на данном листе всего 1 строка (или меньше), т.е. данный лист пуст
        if(rows<=1)
        {
            // освобождение памяти
            usedRange->clear();
            sheet->clear();
            usedRows->clear();
            usedCols->clear();
            sheetNumber++;
            continue;
        }

        data.insert(sheetName, QVariant::fromValue(QStringList()));

        //чтение данных
        for(int row=1; row<=rows; row++)
        {
            QStringList strListRow;
            for(int col=1; col<=cols; col++)
            {
                QScopedPointer<QAxObject> cell (
                            sheet->querySubObject("Cells(QVariant,QVariant)",
                                                  row,
                                                  col)
                            );
                QString result = cell->property("Value").toString();
                strListRow.append(result);
                cell->clear();
            }
            data[sheetName].toStringList().append(strListRow);

            //оставливаем обработку если получено нужное количество строк
            if(_maxCount>0 && row-1 >= _maxCount)
                break;
        }
        sheetNumber++;

        usedRange->clear();
        usedRows->clear();
        usedCols->clear();
        sheet->clear();
    }//end foreach _sheetNames

    sheets->clear();
    workbook->clear();
    workbooks->dynamicCall("Close()");
    workbooks->clear();
    excel->dynamicCall("Quit()");

    qDebug() << "end MyFunction"
             << QThread::currentThread()->currentThreadId()
             << QDateTime::currentDateTime().toString("dd.MM.yyyy hh:mm:ss.zzz");
    return QVariant::fromValue(data);

}

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    qDebug() << "QApplication "
             << QApplication::instance()->thread()->currentThreadId();

    QString str =
            QFileDialog::getOpenFileName(0,
                                         QObject::trUtf8("Укажите исходный файл"),
                                         "",
                                         "Excel (*.xls *.xlsx)");
    if(str.isEmpty())
        qDebug() << "str.isEmpty() ";

    QFuture<QVariant> f1 = QtConcurrent::run(MyFunction,
                                             str,
                                             10);
    f1.waitForFinished();

    qDebug() << "End QApplication "
             << QApplication::instance()->thread()->currentThreadId();
    return app.exec();
}
