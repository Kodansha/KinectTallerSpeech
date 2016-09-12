using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Kinect;
using Microsoft.Speech.AudioFormat;
using Microsoft.Speech.Recognition;


namespace SpeechDemo1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //Esta variable sera la encargada de representar a nuestro dispositivo Kinect y es la que ejecutara
        //algunas de las acciones del hardware
        KinectSensor _sensor;
        //Esta variable speechengine sera quien determine el lenguaje usado, las palabras dictadas,
        //que se va a hacer cuando se reconozca una palabra o frase, entre otros.
        SpeechRecognitionEngine speechengine;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        statusK.Text = "No hay ningun Kinect conectado";
        //En esta linea se indica que cada vez que el estado del dispositivo cambie se mandara llamar al evento KinectSensors_StatusChanged
        KinectSensor.KinectSensors.StatusChanged += new EventHandler<StatusChangedEventArgs>(KinectSensors_StatusChanged);
        //Este metodo conecta activa lo que hara es asignar el primer dispositivo encontrado a nuestra variable _sensor,ademas de inicializarlo e iniciar el reconocimiento de voz
        conectaActiva();

        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Si la variable sensor no es nula, osea que ya fue inicializada, procederemos a detener la entrada de datos de audio y detener el dispositivo
            if (this._sensor != null)
            {
                this._sensor.AudioSource.Stop();
                this._sensor.Stop();
                this._sensor = null;
            }
        }


        //--------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------
        void KinectSensors_StatusChanged(object sender, StatusChangedEventArgs e)
        {
            //Hacemos un switch para ver cual es el estado del dispositivo
            switch (e.Status)
            {
                //En caso de que el status sea Connected quiere decir que hay una conexion correcta entre la PC y el Kinect
                case (KinectStatus.Connected):
                    //De la misma forma mandamos llamar al metodo conectaActiva() el cual inicializara el dispositivo Kinect
                    conectaActiva();
                    break;
                //En caso de que el status sea Disconnected se la variable _sensor se volvera nula e intentaremos buscar otro dispositivo Kinect cuyo estado sea Connected si no se encuentra mandaremos un mensaje indicando que No hay ningun Kinect conectado
                case (KinectStatus.Disconnected):
                    if (this._sensor == e.Sensor)
                    {
                        this._sensor = null;
                        this._sensor = KinectSensor.KinectSensors.FirstOrDefault(x => x.Status == KinectStatus.Connected);
                        if (this._sensor == null)
                        {
                            statusK.Text = "No hay ningun Kinect conectado";
                        }
                    }
                    break;
            }
        }




        //--------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------
        void conectaActiva()
        {
            //Nos aseguramos que la cuenta de sensores conectados sea de al menos 1
            if (KinectSensor.KinectSensors.Count > 0)
            {
                //Checamos que la variable _sensor sea nula
                if (this._sensor == null)
                {
                    //Asignamos el primer sensor Kinect a nuestra variable
                    this._sensor = KinectSensor.KinectSensors[0];
                    if (this._sensor != null)
                    {
                        try
                        {
                            //Iniciamos el dispositivo Kinect
                            this._sensor.Start();
                            //Esto es opcional pero ayuda a colocar el dispositivo Kinect a un cierto angulo de inclinacion, desde -27 a 27
                            _sensor.ElevationAngle = 3;
                            //Informamos que se ha conectado e inicializado correctamente el dispositivo Kinect
                            statusK.Text = "Haz conectado el Kinect";
                        }
                        catch (Exception ex)
                        {
                            //Si hay algun error lo mandamos en el TextBlock statusK
                            statusK.Text = ex.Message.ToString();
                        }

                        //Creamos esta variable ri que tratara de encontrar un language pack valido haciendo uso del metodo obtenerLP
                        RecognizerInfo ri = obtenerLP();
                        //Si se encontro el language pack requerido lo asignaremos a nuestra variable speechengine
                        if (ri != null)
                        {
                            statusK.Text = "Se ha encontrado el languague pack";
                            this.speechengine = new SpeechRecognitionEngine(ri.Id);
                            //Creamos esta variable opciones la cual almacenara las opciones de palabras o frases que podran ser reconocidas por el dispositivo
                            var opciones = new Choices();
                            //Comenzamos a agregar las opciones comenzando por el valor de opcion que tratamos reconocer y una llave que identificara a ese valor
                            //Por ejemplo en esta linea "uno" es el valor de opcion y "UNO" es la llave
                            opciones.Add("uno", "UNO");
                            //En esta linea "unidad" es el valor de opcion y "UNO" es la llave
                            opciones.Add("unidad", "UNO");
                            //En esta linea "dos" es el valor de opcion y "DOS" es la llave
                            opciones.Add("dos", "DOS");
                            //En esta linea "windows ocho" es el valor de opcion y "TRES" es la llave y asi sucesivamente
                            opciones.Add(new SemanticResultValue("windows ocho", "TRES"));
                            opciones.Add(new SemanticResultValue("nuevo windows", "TRES"));

                            //Esta variable creará todo el conjunto de frases y palabras en base a nuestro lenguaje elegido en la variable ri
                            var grammarb = new GrammarBuilder { Culture = ri.Culture };
                            //Agregamos las opciones de palabras y frases a grammarb
                            grammarb.Append(opciones);
                            //Creamos una variable de tipo Grammar utilizando como parametro a grammarb
                            var grammar = new Grammar(grammarb);
                            //Le decimos a nuestra variable speechengine que cargue a grammar
                            this.speechengine.LoadGrammar(grammar);
                            //mandamos llamar al evento SpeechRecognized el cual se ejecutara cada vez que una palabra sea detectada
                            speechengine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(speechengine_SpeechRecognized);
                            //speechengine inicia la entrada de datos de tipo audio
                            speechengine.SetInputToAudioStream(_sensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                            speechengine.RecognizeAsync(RecognizeMode.Multiple);
                        }
                    }
                }
            }
        }

        //--------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------

        private RecognizerInfo obtenerLP()
        {
            //Comienza a checar todos los languagepack que tengamos instalados
            foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                //Aqui es donde elegimos el lenguaje, si se dan cuenta hay una parte donde dice "es-MX" para cambiar el lenguaje a ingles de EU basta con cambiar el valor a "en-US"
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "es-MX".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    //Si se encontro el language pack solicitado se retorna a recognizer
                    return recognizer;
                }
            }
            //En caso de que no se encuentre ningun languaje pack se retorna un valor nulo
            return null;
        }


        //--------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------

        void speechengine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //la variable igualdad sera el porcentaje de igualdad entre la palabra reconocida y el valor de opcion
            //es decir si yo digo "uno" y el valor de opcion es "uno" la igualdad sera mayor al 50 %
            //Si yo digo "jugo" y el valor de opcion es "uno" notaras que el sonido es muy similar pero quizas no mayor al 50 %
            //El valor de porcentaje va de 0.0  a 1.0, ademas notaras que le di un valos de .5 lo cual representa el 50% de igualdad
            const double igualdad = 0.7;
            //Si hay mas del 50% de igualdad con alguna de nuestras opciones
            if (e.Result.Confidence > igualdad)
            {
                Uri src;
                BitmapImage img;
                //haremos un switch para aquellos valores que se componen de unicamente una palabra
                switch (e.Result.Words[0].Text)
                {
                    //En caso de que digamos "uno" la llave "UNO" se abrira y se realizara lo siguiente
                    case "UNO":
                        //Se mandara un mensaje alusivo a la imagen
                        mensaje.Text = "Todo es mas facil con Hotmail. Organiza tu bandeja,usa limpiar para mover o eliminar tus correos automáticamente";
                        mensaje.Background = new SolidColorBrush(Color.FromRgb(247, 126, 5));
                        src = new Uri(@"/Images/img1.png", UriKind.Relative);
                        img = new BitmapImage(src);
                        //El source nuestro control imagen mandara llamar a la imagen 1.jpg , lo mismo se hace para las demas opciones
                        imagen1.Source = img;
                        break;
                    case "DOS":
                        mensaje.Text = "Solo Windows Phone cuenta con un hub de contactos con acceso directo a Facebook, Twitter y Linkedln para que siempre estés al día.";
                        mensaje.Background = new SolidColorBrush(Color.FromRgb(255, 0, 0));
                        src = new Uri(@"/Images/img2.png", UriKind.Relative);
                        img = new BitmapImage(src);
                        imagen1.Source = img;
                        break;
                    default:
                        //En caso de que no solo contenga una palabra tambien realizaremos un switch para ver si la frase corresponde a alguna de nuestros valores de opcion
                        switch (e.Result.Semantics.Value.ToString())
                        {
                            case "TRES":
                                mensaje.Text = "Usando Kinect con speech recognition";
                                mensaje.Background = new SolidColorBrush(Color.FromRgb(5, 134, 247));
                                src = new Uri(@"/Images/img4.jpg", UriKind.Relative);
                                img = new BitmapImage(src);
                                imagen1.Source = img;
                                break;
                            default:
                                mensaje.Text = "No se reconocio el comando";
                                break;
                        }
                        break;

                }
            }
        }

      



        //--------------------------------------------------------------------------------------
        //--------------------------------------------------------------------------------------
    }
}
