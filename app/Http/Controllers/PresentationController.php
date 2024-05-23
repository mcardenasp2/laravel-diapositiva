<?php

namespace App\Http\Controllers;

use GuzzleHttp\Client;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
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

    public function guardarImagen()
    {
        $url = 'https://app.zapinsa.net/storage/uploads/84n3KzrFBC6mRlsprc5dQIua8ju2f1DWUTLuKgcF.png';

        if (filter_var($url, FILTER_VALIDATE_URL) === FALSE) {
            return response()->json(['error' => 'Invalid URL'], 400);
        }

        try {
            $client = new Client([
                'verify' => false,  // Ignorar verificación SSL
            ]);

            $response = $client->get($url);

            if ($response->getStatusCode() !== 200) {
                return response()->json(['error' => 'Failed to download image'], $response->getStatusCode());
            }

            $filename = 'imagen/temp_image_' . time() . '.jpg'; // Guardar en la subcarpeta imagen

            // Crear la subcarpeta si no existe
            if (!Storage::disk('public_folder')->exists('imagen')) {
                Storage::disk('public_folder')->makeDirectory('imagen');
            }

            // Guardar la imagen directamente en la subcarpeta imagen dentro de public
            Storage::disk('public_folder')->put($filename, $response->getBody()->getContents());

            // Construir la URL manualmente
            // $url = env('APP_URL') . '/imagen/' . basename($filename);

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
            $this->addImageToSlide($slide2, public_path($filename) );

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

            // Devolver la URL en la respuesta
            return response()->json(['url' => $url]);
        } catch (\Exception $e) {
            return response()->json(['error' => 'An error occurred while processing the request.'], 500);
        }
    }

    public function guardarImagen3()
    {
        $url = 'https://app.zapinsa.net/storage/uploads/84n3KzrFBC6mRlsprc5dQIua8ju2f1DWUTLuKgcF.png';

        if (filter_var($url, FILTER_VALIDATE_URL) === FALSE) {
            return response()->json(['error' => 'Invalid URL'], 400);
        }

        try {
            $client = new Client([
                'verify' => false,  // Ignorar verificación SSL
            ]);

            $response = $client->get($url);

            if ($response->getStatusCode() !== 200) {
                return response()->json(['error' => 'Failed to download image'], $response->getStatusCode());
            }

            $filename = 'imagen/temp_image_' . time() . '.jpg'; // Guardar en la subcarpeta imagen

            // Crear la subcarpeta si no existe
            if (!Storage::disk('public_folder')->exists('imagen')) {
                Storage::disk('public_folder')->makeDirectory('imagen');
            }

            // Guardar la imagen directamente en la subcarpeta imagen dentro de public
            Storage::disk('public_folder')->put($filename, $response->getBody()->getContents());

            // Construir la URL manualmente
            $url = env('APP_URL') . '/imagen/' . basename($filename);

            // Devolver la URL en la respuesta
            return response()->json(['url' => $url]);
        } catch (\Exception $e) {
            return response()->json(['error' => 'An error occurred while processing the request.'], 500);
        }
    }

    public function guardarImagen2()
    {
        $url = 'https://app.zapinsa.net/storage/uploads/84n3KzrFBC6mRlsprc5dQIua8ju2f1DWUTLuKgcF.png';

        if (filter_var($url, FILTER_VALIDATE_URL) === FALSE) {
            return response()->json(['error' => 'Invalid URL'], 400);
        }

        // $client = new Client();

        $client = new Client([
            'verify' => false,  // Ignorar verificación SSL
        ]);

        $response = $client->get($url);

        if ($response->getStatusCode() !== 200) {
            return response()->json(['error' => 'Failed to download image'], $response->getStatusCode());
        }

        $filename = 'images/temp_image_' . time() . '.jpg';
        $tempPath = storage_path('app/public/' . $filename);

        // Guardar la imagen temporalmente
        Storage::disk('public')->put($filename, $response->getBody()->getContents());

        Storage::disk('public_folder')->put($filename, $response->getBody()->getContents());


        // $url = env('APP_URL') . '/public/' . $filename;
        // dd($filename);
        // $url = env('APP_URL') . '/images/' . basename($filename);
        // dd($url);

        // Descargar la imagen
        $headers = [
            'Content-Type' => 'image/jpeg',
            'Content-Disposition' => 'attachment; filename="' . $filename . '"',
        ];

        $fileContent = Storage::disk('public')->get($filename);

        // $url = Storage::disk('public')->url($filename);
        $url = env('APP_URL') . '/storage/public/' . $filename;
        // dd(public_path('images/imagen.jpg'));
        // dd(public_path($filename) );
        // dd( $url);

        // Eliminar la imagen temporal
        Storage::disk('public')->delete($filename);



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
        $this->addImageToSlide($slide2, public_path($filename) );

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



        return response($fileContent, 200, $headers);
    }
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
