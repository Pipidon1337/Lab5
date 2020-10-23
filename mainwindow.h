#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/qaxobject.h>
#include <ActiveQt/qaxbase.h>
#include <QDebug>
#include <QMessageBox>
#include <QCloseEvent>
#include <QStandardItemModel>
#include <QStandardItem>
#include <QFileDialog>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    QAxObject *mExcel = new QAxObject( "Excel.Application",this);
    QAxObject *workbooks;
    /**
     * @brief row точки
     * @brief col отсчета
     */
    int row = 1;
    int col = 1;
    /**
     * @brief numRows Количество строк
     * @brief numCols Количество столбцов
     */
    int numRows = 100;
    int numCols = 100;
    /**
     * @brief model модель для таблицы
     */
    QStandardItemModel *model = new QStandardItemModel;
    /**
     * @brief item объект для модели
     */
    QStandardItem *item = new QStandardItem(QString(""));
    /**
     * @brief filepath путь к файлу .xlsx
     */
    QString filepath;
    /**
     * @brief cellsList список ячеек
     */
    QList<QVariant> cellsList;
    /**
     * @brief rowsList список строк
     */
    QList<QVariant> rowsList;
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

public slots:
    /**
     * @brief setUp функция начальной установки
     */
    void setUp();
    /**
     * @brief manageXlsxFile функция подготовки файла и пути
     */
    void manageXlsxFile();
    void saveToXlsx();
    /**
     * @brief packToXlsx функция упаковки в файл
     */
    void packToXlsx();
    void getField();
    /**
     * @brief scanData фукнция изъятия данных из таблицы TableView
     */
    void scanData();
    void closeEvent (QCloseEvent *event);
private:
    Ui::MainWindow *ui;
};
#endif // MAINWINDOW_H
