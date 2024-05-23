<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Slide as PhpPresentationSlide;
use PhpOffice\PhpPresentation\Slide\Slide;

class PresentationController extends Controller
{
    // public function createPresentation()
    // {
    //     // Crear una nueva presentación
    //     $presentation = new PhpPresentation();

    //     // Obtener la primera diapositiva
    //     $slide = $presentation->getActiveSlide();

    //     // Crear un cuadro de texto
    //     $shape = $slide->createRichTextShape()
    //                    ->setHeight(300)
    //                    ->setWidth(600)
    //                    ->setOffsetX(170)
    //                    ->setOffsetY(180);

    //     $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
    //     $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

    //     // Añadir texto al cuadro de texto
    //     $textRun = $shape->createTextRun('Hola, esta es una presentación de PowerPoint creada con PHPPresentation en Laravel.');
    //     $textRun->getFont()->setBold(true)
    //                      ->setSize(20)
    //                      ->setColor(new Color('FFFFFFFF'));

    //     // Guardar la presentación como archivo PowerPoint
    //     $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
    //     $fileName = 'presentation.pptx';
    //     $tempFilePath = storage_path($fileName);
    //     $oWriterPPTX->save($tempFilePath);

    //     // Devolver el archivo al usuario
    //     return response()->download($tempFilePath)->deleteFileAfterSend(true);
    // }


    // public function createPresentation()
    // {
    //     // Crear una nueva presentación
    //     $presentation = new PhpPresentation();

    //     // Crear la primera diapositiva
    //     $slide1 = $presentation->getActiveSlide();
    //     $this->addTextToSlide($slide1, 'Esta es la primera diapositiva', 'FF0000FF');

    //     // Crear una segunda diapositiva
    //     $slide2 = $presentation->createSlide();
    //     $this->addTextToSlide($slide2, 'Esta es la segunda diapositiva', 'FF00FF00');

    //     // Crear una tercera diapositiva
    //     $slide3 = $presentation->createSlide();
    //     $this->addTextToSlide($slide3, 'Esta es la tercera diapositiva', 'FFFF0000');

    //     // Guardar la presentación como archivo PowerPoint
    //     $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
    //     $fileName = 'presentation.pptx';
    //     $tempFilePath = storage_path($fileName);
    //     $oWriterPPTX->save($tempFilePath);

    //     // Devolver el archivo al usuario
    //     return response()->download($tempFilePath)->deleteFileAfterSend(true);
    // }

    // private function addTextToSlide(PhpPresentationSlide $slide, $text, $color)
    // {
    //     $shape = $slide->createRichTextShape()
    //                    ->setHeight(300)
    //                    ->setWidth(600)
    //                    ->setOffsetX(170)
    //                    ->setOffsetY(180);
    //     $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
    //     $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

    //     $textRun = $shape->createTextRun($text);
    //     $textRun->getFont()->setBold(true)
    //                      ->setSize(20)
    //                      ->setColor(new Color($color));
    // }


    public function createPresentation()
    {
        // Crear una nueva presentación
        $presentation = new PhpPresentation();

        // Crear la primera diapositiva
        $slide1 = $presentation->getActiveSlide();
        $this->addTextToSlide($slide1, 'Esta es la primera diapositiva con Imagen', 'FF0000FF');
        // dd(asset('images/imagen.jpg'));
        $this->addImageToSlide($slide1, asset('images/imagen.jpg'));

        // Crear una segunda diapositiva
        $slide2 = $presentation->createSlide();
        $ruta = 'https://app.zapinsa.net/storage/uploads/84n3KzrFBC6mRlsprc5dQIua8ju2f1DWUTLuKgcF.png';
        $this->addTextToSlide($slide2, 'Esta es la segunda diapositiva', 'FF00FF00');
        $this->addImageToSlide($slide2, asset($ruta));

        // Crear una tercera diapositiva
        $slide3 = $presentation->createSlide();
        $this->addTextToSlide($slide3, 'Esta es la tercera diapositiva', 'FFFF0000');
        $this->addImageToSlide($slide3, public_path('images/imagen.jpg'));

        // Guardar la presentación como archivo PowerPoint
        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        $fileName = 'presentation.pptx';
        $tempFilePath = storage_path($fileName);
        $oWriterPPTX->save($tempFilePath);

        // Devolver el archivo al usuario
        return response()->download($tempFilePath)->deleteFileAfterSend(true);
    }

    private function addTextToSlide(PhpPresentationSlide $slide, $text, $color)
    {
        $shape = $slide->createRichTextShape()
                       ->setHeight(300)
                       ->setWidth(600)
                       ->setOffsetX(170)
                       ->setOffsetY(180);
        $shape->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $textRun = $shape->createTextRun($text);
        $textRun->getFont()->setBold(true)
                         ->setSize(20)
                         ->setColor(new Color($color));
    }

    private function addImageToSlide(PhpPresentationSlide $slide, $imagePath)
    {
        if (!file_exists($imagePath)) {
            return;
        }

        $shape = $slide->createDrawingShape();
        $shape->setName('Imagen')
              ->setDescription('Descripción de la imagen')
              ->setPath($imagePath)
              ->setHeight(200)
              ->setOffsetX(170)
              ->setOffsetY(400);
    }
}
